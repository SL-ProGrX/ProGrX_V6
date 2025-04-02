VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpspr80.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.2#0"; "Codejock.Controls.v19.2.0.ocx"
Begin VB.Form frmCR_CatalogoDestinos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Destinos de las Líneas de Crédito"
   ClientHeight    =   7212
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   12564
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7212
   ScaleWidth      =   12564
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   5892
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   12372
      _Version        =   1245186
      _ExtentX        =   21823
      _ExtentY        =   10393
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
      Item(0).Caption =   "Destinos"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "vGrid"
      Item(1).Caption =   "Asignación"
      Item(1).ControlCount=   6
      Item(1).Control(0)=   "lsw"
      Item(1).Control(1)=   "lswDes"
      Item(1).Control(2)=   "rbCatalogo(0)"
      Item(1).Control(3)=   "rbCatalogo(1)"
      Item(1).Control(4)=   "lbl"
      Item(1).Control(5)=   "cmdReporte"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   4332
         Left            =   -69880
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   5892
         _Version        =   1245186
         _ExtentX        =   10393
         _ExtentY        =   7641
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
         HotTracking     =   -1  'True
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.ListView lswDes 
         Height          =   4332
         Left            =   -63880
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   6012
         _Version        =   1245186
         _ExtentX        =   10604
         _ExtentY        =   7641
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
         HotTracking     =   -1  'True
         Appearance      =   16
         UseVisualStyle  =   0   'False
      End
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   372
         Left            =   -59200
         TabIndex        =   8
         Top             =   5280
         Visible         =   0   'False
         Width           =   1332
         _Version        =   1245186
         _ExtentX        =   2350
         _ExtentY        =   656
         _StockProps     =   79
         Caption         =   "Informe"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   5292
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   12012
         _Version        =   524288
         _ExtentX        =   21188
         _ExtentY        =   9334
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
         MaxCols         =   497
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_CatalogoDestinos.frx":0000
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.RadioButton rbCatalogo 
         Height          =   252
         Index           =   0
         Left            =   -68320
         TabIndex        =   5
         Top             =   5280
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1245186
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Créditos"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
         Value           =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton rbCatalogo 
         Height          =   252
         Index           =   1
         Left            =   -66400
         TabIndex        =   6
         Top             =   5280
         Visible         =   0   'False
         Width           =   1572
         _Version        =   1245186
         _ExtentX        =   2773
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Retenciones"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   16
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
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
         Height          =   312
         Left            =   -69880
         TabIndex        =   7
         Top             =   480
         Visible         =   0   'False
         Width           =   12012
      End
   End
   Begin XtremeSuiteControls.PushButton cmdActualiza 
      Height          =   372
      Left            =   11160
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   492
      _Version        =   1245186
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin XtremeSuiteControls.PushButton cmdModifica 
      Height          =   372
      Left            =   11640
      TabIndex        =   10
      Top             =   720
      Visible         =   0   'False
      Width           =   492
      _Version        =   1245186
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Destinos o Plan de Inversión del Préstamo"
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
      Height          =   480
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   300
      Width           =   7692
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   12732
   End
End
Attribute VB_Name = "frmCR_CatalogoDestinos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String, vPaso As Boolean

Private Sub cmdReporte_Click()
With frmContenedor.Crt
   .Reset
   .WindowShowGroupTree = True
   .WindowShowPrintSetupBtn = True
   .WindowShowRefreshBtn = True
   .WindowShowSearchBtn = True
   .WindowState = crptMaximized
   .WindowTitle = "Reportes del Módulo de Crédito"

   .Connect = glogon.ConectRPT

   .Formulas(0) = "Empresa = '" & GLOBALES.gstrNombreEmpresa & "'"
   .ReportFileName = SIFGlobal.fxPathReportes("Credito_CatalogoDestinos.rpt")
   .PrintReport
End With

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3

tcMain.Item(0).Selected = True

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
 
strSQL = "select cod_destino,descripcion,tasa,tbp,int_form,case when isnull(TCIntForma,'A') = 'A' then 'Adelantado' else 'Vencido' end as 'TipoCbrInt'" _
       & ", primer_cuota,ENVIO_TESORERIA,prioridad from Catalogo_Destinos" _
       & " order by cod_destino"
Call sbCargaGrid(vGrid, 9, strSQL)


With lsw.ColumnHeaders
  .Clear
  .Add , , "Código", 1440
  .Add , , "Descripción", 3600
End With


With lswDes.ColumnHeaders
  .Clear
  .Add , , "Código", 1440
  .Add , , "Descripción", 3600
End With


Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.Enabled = cmdModifica.Enabled
lswDes.Enabled = cmdActualiza.Enabled
End Sub

Private Function fxExiste(vCod As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from Catalogo_Destinos" _
       & " where cod_destino = '" & vCod & "'"
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

'int_form,primer_cuota,ENVIO_TESORERIA,prioridad

If vGrid.Text = "" Then Exit Function


If Not fxExiste(vGrid.Text) Then
   vGrid.col = 1
   strSQL = "insert Catalogo_Destinos(cod_destino,descripcion,tasa,TBP,int_form, TCIntForma, primer_cuota,envio_tesoreria,prioridad)" _
          & " values('" & vGrid.Text & "','"
   vGrid.col = 2
   strSQL = strSQL & vGrid.Text & "',"
   vGrid.col = 3
   strSQL = strSQL & CCur(vGrid.Text) & ","
   vGrid.col = 4
   strSQL = strSQL & vGrid.Value & ","
   vGrid.col = 5
   strSQL = strSQL & vGrid.Value & ",'"
   vGrid.col = 6
   strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "',"
   vGrid.col = 7
   strSQL = strSQL & vGrid.Value & ","
   vGrid.col = 8
   strSQL = strSQL & vGrid.Value & ",'"
   vGrid.col = 9
   strSQL = strSQL & vGrid.Text & "')"
   
   Call ConectionExecute(strSQL)
   vGrid.col = 1
   Call Bitacora("Registra", "Destino de Linea Cod: " & vGrid.Text)
   
 Else 'Actualizar
    vGrid.col = 2
    strSQL = "update Catalogo_Destinos set descripcion = '" & vGrid.Text
    vGrid.col = 3
    strSQL = strSQL & "',tasa= " & CCur(vGrid.Text) & ",Tbp = "
    vGrid.col = 4
    strSQL = strSQL & vGrid.Value & ",int_form = "
    vGrid.col = 5
    strSQL = strSQL & vGrid.Value & ",TCIntForma = '"
    vGrid.col = 6
    strSQL = strSQL & Mid(vGrid.Text, 1, 1) & "', primer_cuota = "
    vGrid.col = 7
    strSQL = strSQL & vGrid.Value & ",envio_tesoreria = "
    vGrid.col = 8
    strSQL = strSQL & vGrid.Value & ",prioridad = '"
    vGrid.col = 9
    strSQL = strSQL & vGrid.Text & "'"
    
    vGrid.col = 1
    strSQL = strSQL & " where cod_destino = '" & vGrid.Text & "'"
   
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Destino Linea Cod: " & vGrid.Text)
    
End If

Exit Function
   
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

   
End Function

Private Sub sbCargaLswAdicional()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

strSQL = "select R.*,A.codigo as Existe" _
       & " from Catalogo_Destinos R left Join Catalogo_DestinosAsg A " _
       & " on R.cod_destino = A.cod_destino and A.codigo = '" _
       & vCodigo & "' order by existe desc,R.cod_destino"
Call OpenRecordSet(rs, strSQL, 0)

vPaso = True
lswDes.ListItems.Clear

Do While Not rs.EOF
  Set itmX = lswDes.ListItems.Add(, , rs!cod_destino)
      itmX.SubItems(1) = rs!Descripcion & ""
      itmX.Checked = IIf(IsNull(rs!Existe), False, True)
      
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

  vCodigo = Item.Text
  lbl.Caption = Item.Text & " - " & Item.SubItems(1)
  Call sbCargaLswAdicional

End Sub


Private Sub lswDes_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lswDes.SortKey = ColumnHeader.Index - 1
  If lswDes.SortOrder = 0 Then lswDes.SortOrder = 1 Else lswDes.SortOrder = 0
  lswDes.Sorted = True
End Sub

Private Sub lswDes_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError


If Item.Checked Then
    strSQL = "insert Catalogo_DestinosAsg(cod_destino,codigo) values('" _
           & Item.Text & "','" & vCodigo & "')"
Else
    strSQL = "delete Catalogo_DestinosAsg where cod_destino = '" _
           & Item.Text & "' and codigo = '" & vCodigo & "'"
End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub rbCatalogo_Click(Index As Integer)

Call sbCargaLineas

End Sub



Private Sub sbCargaLineas()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

If tcMain.Item(0).Selected = True Then Exit Sub

Me.MousePointer = vbHourglass

vCodigo = ""
lbl.Caption = ""
lswDes.ListItems.Clear

Select Case True
  Case rbCatalogo.Item(0).Value
    strSQL = "select codigo,descripcion" _
           & " from catalogo where (Retencion = 'N' and Poliza = 'N')" _
           & " order by codigo"
  Case rbCatalogo.Item(1).Value
    strSQL = "select codigo,descripcion" _
           & " from catalogo where (Retencion = 'S' or Poliza = 'S')" _
           & " order by codigo"
End Select

Call OpenRecordSet(rs, strSQL, 0)

vPaso = True
    lsw.ListItems.Clear
    Do While Not rs.EOF
     Set itmX = lsw.ListItems.Add(, , rs!Codigo)
         itmX.SubItems(1) = rs!Descripcion & ""
     rs.MoveNext
    Loop
    rs.Close
vPaso = False

Me.MousePointer = vbDefault
End Sub

Private Sub tcMain_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

If Item.Index = 0 Then Exit Sub
Call sbCargaLineas

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



