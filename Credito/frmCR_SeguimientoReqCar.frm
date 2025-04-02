VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#19.1#0"; "Codejock.Controls.v19.1.0.ocx"
Begin VB.Form frmCR_SeguimientoReqCar 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "..."
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10005
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   16.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.ListView lsw 
      Height          =   2652
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   9732
      _Version        =   1245185
      _ExtentX        =   17166
      _ExtentY        =   4678
      _StockProps     =   77
      ForeColor       =   -2147483630
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
   End
   Begin FPSpreadADO.fpSpread vGrid 
      Height          =   6252
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   9732
      _Version        =   524288
      _ExtentX        =   17166
      _ExtentY        =   11028
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
      SpreadDesigner  =   "frmCR_SeguimientoReqCar.frx":0000
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.RadioButton optTipo 
      Height          =   252
      Index           =   0
      Left            =   5400
      TabIndex        =   4
      Top             =   1440
      Width           =   1932
      _Version        =   1245185
      _ExtentX        =   3408
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cargos Manuales"
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
      Value           =   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   9240
      Top             =   6600
   End
   Begin FPSpreadADO.fpSpread vGridPrima 
      Height          =   2652
      Left            =   2280
      TabIndex        =   2
      Top             =   4560
      Width           =   7452
      _Version        =   524288
      _ExtentX        =   13145
      _ExtentY        =   4678
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
      MaxCols         =   495
      ScrollBars      =   2
      SpreadDesigner  =   "frmCR_SeguimientoReqCar.frx":05CF
      VScrollSpecialType=   2
      AppearanceStyle =   1
   End
   Begin XtremeSuiteControls.RadioButton optTipo 
      Height          =   252
      Index           =   1
      Left            =   7560
      TabIndex        =   5
      Top             =   1440
      Width           =   2292
      _Version        =   1245185
      _ExtentX        =   4043
      _ExtentY        =   444
      _StockProps     =   79
      Caption         =   "Cargos Registrados"
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
      Transparent     =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   16
   End
   Begin VB.Label Label1 
      Caption         =   "Prima de recargo por adquisición de la operación."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1332
      Left            =   600
      TabIndex        =   3
      Top             =   4920
      Width           =   1332
   End
   Begin VB.Label lblX 
      BackStyle       =   0  'Transparent
      Caption         =   "..."
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
      Height          =   492
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   8052
   End
   Begin VB.Image imgBanner 
      Height          =   1215
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmCR_SeguimientoReqCar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean

Private Sub Form_Load()

vModulo = 3

vGrid.AppearanceStyle = fxGridStyle
Set imgBanner.Picture = frmContenedor.imgBanner_Tramites.Picture

With lsw.ColumnHeaders
  .Clear
  .Add , , "Código", 900
  .Add , , "Descripción", 2900
  .Add , , "Monto", 1200, vbRightJustify
  .Add , , "Tipo", 1200
  .Add , , "Valor", 1200, vbRightJustify
  .Add , , "PlazoTipo", 1200, vbRightJustify
  .Add , , "PlazoDias", 1200, vbRightJustify
  .Add , , "Diferido?", 1200, vbCenter
End With


'Si esta anulada o formalizada, no permitir modificaciones
If Operacion.EstadoSolicitud = "N" Or Operacion.EstadoSolicitud = "F" Then
   vGrid.Enabled = False
   lsw.Enabled = False
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer, strSQL As String
Dim vCodReq As String, vEstado As String


If Operacion.Ventana = "R" And vGrid.Enabled Then
  strSQL = "delete operacion_requisitos where id_solicitud = " & Operacion.Operacion _
         & " and cod_requisito in(select cod_requisito from requisitos_adicionales where visible = 1)"
  
  For i = 1 To vGrid.MaxRows
    vGrid.Row = i
    
    vGrid.col = 2
    vCodReq = vGrid.Text
    
    vGrid.col = 1
    Select Case vGrid.Value
       Case 0
         vEstado = "P" 'Pendiente
       Case 1
         vEstado = "M" 'Marcado
       Case 2
         vEstado = "D" 'Desmarcado
    End Select
    strSQL = strSQL & Space(10) & "insert operacion_requisitos(cod_requisito,id_solicitud,codigo,estado,opcional) values('" _
           & vCodReq & "'," & Operacion.Operacion & ",'" & Operacion.Codigo & "'," & vGrid.Value & "," & vGrid.CellTag & ")"
  Next i

 'Procesa Lote
 Call ConectionExecute(strSQL)


 'Revisa Requisitos no Visibles
 strSQL = "exec spCrdRequisitosOperacionNoVisiblesEstado " & Operacion.Operacion
 Call ConectionExecute(strSQL)

End If

End Sub



Private Sub lsw_ItemCheck(ByVal Item As XtremeSuiteControls.ListViewItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   
   strSQL = "insert operacion_cargos(cod_cargo,id_solicitud,codigo,tipo,monto,valor,plazo_tipo,plazo_dias" _
          & ",tipo_deduccion,diferido) values('" & Item.Text & "'," & Operacion.Operacion & ",'" & Operacion.Codigo _
          & "','" & Mid(Item.SubItems(3), 1, 1) & "'," & CCur(Item.SubItems(2)) & "," & CDbl(Item.SubItems(4)) & ",'" & Item.SubItems(5) & "'," _
          & Item.SubItems(6) & ",'F'," & Item.SubItems(7) & ")"
Else
   strSQL = "delete operacion_cargos where id_solicitud = " & Operacion.Operacion & " and cod_cargo = '" _
          & Item.Text & "'"

End If
Call ConectionExecute(strSQL)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub optTipo_Click(Index As Integer)
Call Timer1_Timer
End Sub

Private Sub Timer1_Timer()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListViewItem

Me.MousePointer = vbHourglass

Timer1.Interval = 0

vPaso = True

lsw.ListItems.Clear

lsw.Visible = False
vGrid.Visible = False

If Operacion.Ventana = "R" Then
  vGrid.Visible = True
  vGrid.MaxRows = 0
  
'  vGrid.Top = 960
'  vGrid.Left = 240
'
  Me.Caption = "REQUISITOS - Indique el cumplimiento de requisitos"
  lblX.Caption = "Requisitos"

  strSQL = "exec spCrdRequisitosOperacionLista " & Operacion.Operacion
  Call OpenRecordSet(rs, strSQL)
  
  If rs.EOF And rs.BOF Then
     rs.Close
     Exit Sub
  End If
  
  vGrid.MaxCols = 3
  vGrid.MaxRows = 1
  Do While Not rs.EOF
    vGrid.Row = vGrid.MaxRows
    vGrid.col = 2
    vGrid.Text = CStr(rs!COD_REQUISITO)
    vGrid.col = 3
    vGrid.Text = CStr(rs!Descripcion)
        
    vGrid.col = 1
    vGrid.CellTag = rs!Opcional
    vGrid.Value = rs!Estado

    vGrid.MaxRows = vGrid.MaxRows + 1
    rs.MoveNext
  Loop
  rs.Close
  
  vGrid.MaxRows = vGrid.MaxRows - 1

Else

  Me.Caption = "CARGOS ADICIONALES"
  lblX.Caption = "Cargos Adicionales"
  
  lsw.Checkboxes = False
  lsw.Visible = True

  If optTipo.Item(0).Value = True Then
    'Cargos Manuales
    lsw.Checkboxes = True
    strSQL = "select R.cod_cargo,R.descripcion,A.tipo,A.valor,A.monto,R.plazo_tipo,R.plazo_dias,R.Diferido_Cargo" _
           & ",dbo.fxCRDOperacionCargoValor(A.id_solicitud,R.cod_cargo,dbo.MyGetdate()) as Cargo" _
           & " from cargos_adicionales R inner join operacion_cargos A" _
           & " on R.cod_cargo = A.cod_cargo and A.id_solicitud = " & Operacion.Operacion _
           & " and R.Automatico = 0 and R.Base in ('C','A')"
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!COD_CARGO)
          itmX.SubItems(1) = rs!Descripcion
          itmX.SubItems(2) = Format(rs!Cargo, "Standard")
          itmX.SubItems(3) = IIf((rs!Tipo = "P"), "Porcentaje", "Monto")
          itmX.SubItems(4) = Format(rs!Valor, "Standard")
          itmX.SubItems(5) = Format(rs!plazo_Tipo, "Standard")
          itmX.SubItems(6) = Format(rs!plazo_dias, "Standard")
          itmX.SubItems(7) = rs!Diferido_Cargo
          
          itmX.Checked = True
      rs.MoveNext
    Loop
    rs.Close
    
    'Paso 2: Llena cargos no asignados (Manuales, x Linea x Destino x Garantia)
    strSQL = "select cod_destino, garantia from reg_creditos" _
           & " where id_solicitud = " & Operacion.Operacion
    Call OpenRecordSet(rs, strSQL)
    
    strSQL = "select *,dbo.fxCRDOperacionCargoValor(" & Operacion.Operacion & ",cod_cargo,dbo.MyGetdate()) as Cargo" _
           & " from cargos_adicionales " _
           & " where cod_cargo in(select cod_cargo from CRD_CARGOS_ASG_DETALLE where codigo = '" & Operacion.Codigo _
           & "' and cod_destino = '" & rs!cod_destino & "' and Garantia = '" & rs!Garantia _
           & "' and cod_cargo not in(select cod_cargo from operacion_cargos where id_solicitud = " & Operacion.Operacion & "))" _
           & "  and Automatico = 0 and Base in ('C','A')"
    rs.Close
    Call OpenRecordSet(rs, strSQL, 0)
    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!COD_CARGO)
          itmX.SubItems(1) = rs!Descripcion
          itmX.SubItems(2) = Format(rs!Cargo, "Standard")
          itmX.SubItems(3) = IIf((rs!Tipo = "P"), "Porcentaje", "Monto")
          itmX.SubItems(4) = Format(rs!Valor, "Standard")
          itmX.SubItems(5) = Format(rs!plazo_Tipo, "Standard")
          itmX.SubItems(6) = Format(rs!plazo_dias, "Standard")
          itmX.SubItems(7) = rs!Diferido_Cargo
          itmX.Checked = False
      rs.MoveNext
    Loop
    rs.Close
    
    
  Else
    'Cargos Asignados
    strSQL = "select R.cod_cargo,R.descripcion,A.tipo,A.valor,A.monto,R.plazo_tipo,R.plazo_dias" _
           & " from cargos_adicionales R inner join operacion_cargos A" _
           & " on R.cod_cargo = A.cod_cargo and A.id_solicitud = " & Operacion.Operacion
    Call OpenRecordSet(rs, strSQL)
    Do While Not rs.EOF
      Set itmX = lsw.ListItems.Add(, , rs!COD_CARGO)
          itmX.SubItems(1) = rs!Descripcion
          itmX.SubItems(2) = Format(rs!Monto, "Standard")
          itmX.SubItems(3) = IIf((rs!Tipo = "P"), "Porcentaje", "Monto")
          itmX.SubItems(4) = Format(rs!Valor, "Standard")
          itmX.SubItems(5) = Format(rs!plazo_Tipo, "Standard")
          itmX.SubItems(6) = Format(rs!plazo_dias, "Standard")
      rs.MoveNext
    Loop
    rs.Close
    
   End If 'Cargos Manuales / Asignados
   
   'Carga Primas
   strSQL = "select R.cod_cargo,R.descripcion,isnull(A.monto,0) as 'Monto'" _
           & " from cargos_adicionales R left join operacion_cargos A" _
           & " on R.cod_cargo = A.cod_cargo and A.id_solicitud = " & Operacion.Operacion _
           & " Where R.Base = 'P'"
   Call sbCargaGrid(vGridPrima, 3, strSQL)
   vGridPrima.MaxRows = vGridPrima.MaxRows - 1
   
End If 'Requisitos / Cargos

vPaso = False

Me.MousePointer = vbDefault

End Sub



Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCuenta As String, vTipo As String

On Error GoTo vError

fxGuardar = 0

With vGridPrima

    .Row = .ActiveRow
    .col = 1
    
    strSQL = "exec spCrdOperacionFormalizaPrima " & Operacion.Operacion & ",'" & .Text & "',"
    .col = 3
    strSQL = strSQL & CCur(.Text) & ",'" & glogon.Usuario & "'"
    
    Call ConectionExecute(strSQL)

End With
fxGuardar = 1

Exit Function

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Function


Private Sub vGridPrima_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, i As Long


On Error GoTo vError

If vGridPrima.ActiveCol = vGridPrima.MaxCols And (KeyCode = vbKeyReturn Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  vGridPrima.Row = vGridPrima.ActiveRow
  vGridPrima.col = 1
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
