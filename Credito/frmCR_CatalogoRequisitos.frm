VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmCR_CatalogoRequisitos 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requisitos de Formalización de Operaciones de Créditos"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   Icon            =   "frmCR_CatalogoRequisitos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   9045
   Begin XtremeSuiteControls.TabControl tcMain 
      Height          =   6852
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   8772
      _Version        =   1441793
      _ExtentX        =   15473
      _ExtentY        =   12086
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
      SelectedItem    =   1
      Item(0).Caption =   "Requisitos"
      Item(0).ControlCount=   2
      Item(0).Control(0)=   "vGrid"
      Item(0).Control(1)=   "cmdReporte"
      Item(1).Caption =   "Asignación"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "scTitulo"
      Item(1).Control(1)=   "cboRequisitos"
      Item(1).Control(2)=   "Label6(0)"
      Item(1).Control(3)=   "vGridAsg"
      Item(1).Control(4)=   "lsw"
      Begin XtremeSuiteControls.ListView lsw 
         Height          =   2292
         Left            =   360
         TabIndex        =   10
         Top             =   840
         Width           =   8052
         _Version        =   1441793
         _ExtentX        =   14203
         _ExtentY        =   4043
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
      End
      Begin FPSpreadADO.fpSpread vGrid 
         Height          =   6012
         Left            =   -69760
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   8292
         _Version        =   524288
         _ExtentX        =   14626
         _ExtentY        =   10605
         _StockProps     =   64
         AllowCellOverflow=   -1  'True
         ArrowsExitEditMode=   -1  'True
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
         SpreadDesigner  =   "frmCR_CatalogoRequisitos.frx":6852
         VScrollSpecial  =   -1  'True
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.ComboBox cboRequisitos 
         Height          =   312
         Left            =   6240
         TabIndex        =   6
         Top             =   396
         Width           =   2172
         _Version        =   1441793
         _ExtentX        =   3836
         _ExtentY        =   582
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
      Begin FPSpreadADO.fpSpread vGridAsg 
         Height          =   3012
         Left            =   360
         TabIndex        =   8
         Top             =   3600
         Width           =   8052
         _Version        =   524288
         _ExtentX        =   14203
         _ExtentY        =   5313
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         DisplayRowHeaders=   0   'False
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
         MaxCols         =   4
         ScrollBars      =   2
         SpreadDesigner  =   "frmCR_CatalogoRequisitos.frx":6E65
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin XtremeSuiteControls.PushButton cmdReporte 
         Height          =   300
         Left            =   -62680
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   972
         _Version        =   1441793
         _ExtentX        =   1714
         _ExtentY        =   529
         _StockProps     =   79
         Caption         =   "Informe"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   6
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   252
         Index           =   0
         Left            =   4440
         TabIndex        =   7
         Top             =   396
         Width           =   1692
         _Version        =   1441793
         _ExtentX        =   2984
         _ExtentY        =   444
         _StockProps     =   79
         Caption         =   "Nivel de aplicación:"
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   1
         Transparent     =   -1  'True
      End
      Begin XtremeShortcutBar.ShortcutCaption scTitulo 
         Height          =   372
         Left            =   360
         TabIndex        =   5
         Top             =   3120
         Width           =   8052
         _Version        =   1441793
         _ExtentX        =   14203
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
         VisualTheme     =   3
      End
   End
   Begin XtremeSuiteControls.PushButton cmdActualiza 
      Height          =   372
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "..."
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
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
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   492
      _Version        =   1441793
      _ExtentX        =   868
      _ExtentY        =   656
      _StockProps     =   79
      Caption         =   "..."
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   7.5
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
      Caption         =   "Requisitos de Créditos"
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
      Index           =   2
      Left            =   2160
      TabIndex        =   0
      Top             =   300
      Width           =   5772
   End
   Begin VB.Image imgBanner 
      Height          =   1092
      Left            =   0
      Top             =   0
      Width           =   10812
   End
End
Attribute VB_Name = "frmCR_CatalogoRequisitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub cboRequisitos_Click()
If vPaso Then Exit Sub
Call sbLista_Consulta
End Sub

Private Sub cmdReporte_Click()

Me.MousePointer = vbHourglass

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
   .ReportFileName = SIFGlobal.fxPathReportes("Credito_CatalogoRequisitos.rpt")
   .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 3

Set imgBanner.Picture = frmContenedor.imgBanner_Mantenimiento.Picture
vGrid.AppearanceStyle = fxGridStyle

tcMain.Item(0).Selected = True

With lsw.ColumnHeaders
    .Clear
    .Add , , "Código", 1500, vbCenter
    .Add , , "Descripción", 6450
End With

vPaso = True
    strSQL = "select cod_requisito,descripcion,visible,0 from requisitos_adicionales" _
           & " order by cod_requisito"
    Call sbCargaGrid(vGrid, 3, strSQL)

    cboRequisitos.Clear
    cboRequisitos.AddItem "Línea"
    cboRequisitos.AddItem "Garantía"
    cboRequisitos.Text = "Garantía"
vPaso = False

scTitulo.Caption = ""
scTitulo.Tag = ""

Call cboRequisitos_Click

Call Formularios(Me)
Call RefrescaTags(Me)

vGrid.Enabled = cmdModifica.Enabled
vGridAsg.Enabled = vGrid.Enabled
End Sub

Private Function fxExiste(vCod As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select isnull(count(*),0) as Existe from requisitos_adicionales" _
       & " where cod_requisito = '" & vCod & "'"
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
   strSQL = "insert requisitos_adicionales(cod_requisito,descripcion,visible)" _
          & " values('" & vGrid.Text & "','"
   vGrid.col = 2
   strSQL = strSQL & vGrid.Text & "',"
   vGrid.col = 3
   strSQL = strSQL & vGrid.Value & ")"
   Call ConectionExecute(strSQL)
   vGrid.col = 1
   Call Bitacora("Registra", "Requisito Adicional Cod: " & vGrid.Text)
   
 Else 'Actualizar
    vGrid.col = 2
    strSQL = "update requisitos_adicionales set descripcion = '" & vGrid.Text
    vGrid.col = 3
    strSQL = strSQL & "', visible = " & vGrid.Value
    vGrid.col = 1
    strSQL = strSQL & " where cod_requisito = '" & vGrid.Text & "'"
   
    Call ConectionExecute(strSQL)
    
    Call Bitacora("Modifica", "Requisito Adicional Cod: " & vGrid.Text)
    
End If

Exit Function
   
vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

   
End Function

Private Sub sbCargaAsignacion()
Dim strSQL As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

With vGridAsg
    .MaxRows = 0
    .MaxCols = 4
    
    vPaso = True

    Select Case Mid(cboRequisitos.Text, 1, 1)
      
      Case "L" 'Carga Requisitos x Linea
            strSQL = "select R.*,isnull(A.opcional,0) as 'OpcionalX',A.codigo as Existe" _
                   & " from Requisitos_Adicionales R left Join Requisitos_asignacion A " _
                   & " on R.cod_requisito = A.cod_requisito and A.codigo = '" _
                   & scTitulo.Tag & "' order by existe desc,R.cod_requisito"
            
      Case "G" 'Carga Requisitos x Garantía
            strSQL = "select R.*,isnull(A.opcional,0) as 'OpcionalX',A.Garantia as Existe" _
                   & " from Requisitos_Adicionales R left Join CRD_GARANTIA_REQUISITOS A " _
                   & " on R.cod_requisito = A.cod_requisito and A.Garantia = '" _
                   & scTitulo.Tag & "' order by existe desc,R.cod_requisito"
    End Select 'Nivel del Requisito
    
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
      .MaxRows = .MaxRows + 1
      .Row = .MaxRows
      .col = 1
      .Text = rs!COD_REQUISITO
      .col = 2
      .Text = rs!Descripcion
      .col = 3
      .Value = rs!OpcionalX
      .col = 4
      .Value = IIf(IsNull(rs!Existe), 0, 1)
      rs.MoveNext
    Loop
    rs.Close



vPaso = False

End With

Me.MousePointer = vbDefault

End Sub

'Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
'Dim strSQL As String
'
'If vPaso Then Exit Sub
'
'On Error GoTo vError
'
'If Item.Checked Then
'    strSQL = "insert requisitos_asignacion(cod_requisito,codigo) values('" _
'           & Item.Text & "','" & scTitulo.Tag & "')"
'Else
'    strSQL = "delete requisitos_asignacion where cod_requisito = '" _
'           & Item.Text & "' and codigo = '" & scTitulo.Tag & "'"
'End If
'Call ConectionExecute(strSQL)
'
'Exit Sub
'
'vError:
' MsgBox fxSys_Error_Handler(Err.Description), vbCritical
'
'End Sub

Private Sub sbLista_Consulta()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem


If vPaso Then Exit Sub
If tcMain.Selected.Index = 0 Then Exit Sub

Me.MousePointer = vbHourglass

scTitulo.Caption = ""
scTitulo.Tag = ""

vGridAsg.MaxRows = 0

If Mid(cboRequisitos.Text, 1, 1) = "L" Then
    strSQL = "select codigo,descripcion from catalogo where retencion = 'N' and poliza = 'N' and requisitos_tipo = 'L' order by codigo"
Else
    strSQL = "select garantia as Codigo,descripcion from Crd_Garantia_Tipos order by Garantia"
End If

Call OpenRecordSet(rs, strSQL, 0)

lsw.ListItems.Clear
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!Codigo)
     itmX.SubItems(1) = rs!Descripcion & ""
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub


Private Sub lsw_ColumnClick(ByVal ColumnHeader As XtremeSuiteControls.ListViewColumnHeader)
 lsw.SortKey = ColumnHeader.Index - 1
  If lsw.SortOrder = 0 Then lsw.SortOrder = 1 Else lsw.SortOrder = 0
  lsw.Sorted = True
End Sub

Private Sub lsw_ItemClick(ByVal Item As XtremeSuiteControls.ListViewItem)
  
If vPaso Then Exit Sub
  
  scTitulo.Tag = lsw.SelectedItem.Text
  scTitulo.Caption = "[" & Item.Text & "] " & Item.SubItems(1)
  
  Call sbCargaAsignacion

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

Private Sub vGridAsg_ButtonClicked(ByVal col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
Dim strSQL As String, vMovimiento As String
Dim vTempo As Integer


If vPaso Then Exit Sub


With vGridAsg

     .Row = Row
     .col = col
     
      
     
Select Case Mid(cboRequisitos.Text, 1, 1)
  Case "L" 'Requisitos a Nivel de Línea
     If col = 4 Then 'Ultima Columna
        If .Value = 1 Then
           .col = 3
           vTempo = .Value
           .col = 1
           vMovimiento = "Registra"
           strSQL = "insert requisitos_asignacion(codigo,cod_requisito,opcional) values('" _
                  & scTitulo.Tag & "','" & .Text & "'," & vTempo & ")"
        Else
           .col = 1
           vMovimiento = "Borrar"
           strSQL = "delete requisitos_asignacion where codigo = '" _
                  & scTitulo.Tag & "' and cod_requisito = '" & .Text & "'"
           
         End If
         
         Call ConectionExecute(strSQL)
         Call Bitacora(vMovimiento, "Requisito : " & .Text & " a la Línea: " & scTitulo.Tag)
     End If
  
     If col = 3 Then 'Columna de Opcional
        .col = 3
        vTempo = .Value
        .col = 4
        If .Value = 1 Then
            .col = 1
            vMovimiento = "Modifica"
            strSQL = "update requisitos_asignacion set Opcional = " & vTempo & " where codigo = '" _
                   & scTitulo.Tag & "' and cod_requisito = '" & .Text & "'"
            
            Call ConectionExecute(strSQL)
            Call Bitacora(vMovimiento, "Requisito : " & .Text & " a la Línea: " & scTitulo.Tag)
        End If
     End If
  
  Case "G" 'Requisitos a Nivel de Garantía
     If col = 4 Then 'Ultima Columna
        If .Value = 1 Then
           .col = 3
           vTempo = .Value
           .col = 1
           vMovimiento = "Registra"
           strSQL = "insert CRD_GARANTIA_REQUISITOS(garantia,cod_requisito,opcional) values('" _
                  & scTitulo.Tag & "','" & .Text & "'," & vTempo & ")"
        Else
           .col = 1
           vMovimiento = "Borrar"
           strSQL = "delete CRD_GARANTIA_REQUISITOS where garantia = '" _
                  & scTitulo.Tag & "' and cod_requisito = '" & .Text & "'"
           
         End If
         
         Call ConectionExecute(strSQL)
         Call Bitacora(vMovimiento, "Requisito : " & .Text & " a la Garantía: " & scTitulo.Tag)
     End If
  
     If col = 3 Then 'Columna de Opcional
        .col = 3
        vTempo = .Value
        .col = 4
        If .Value = 1 Then
            .col = 1
            vMovimiento = "Modifica"
            strSQL = "update CRD_GARANTIA_REQUISITOS set Opcional = " & vTempo & " where garantia = '" _
                   & scTitulo.Tag & "' and cod_requisito = '" & .Text & "'"
            
            Call ConectionExecute(strSQL)
            Call Bitacora(vMovimiento, "Requisito : " & .Text & " a la Garantía: " & scTitulo.Tag)
        End If
     End If
     
End Select


End With


End Sub
