VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCO_UnificacionCuotas 
   Caption         =   "Unificación de Cuotas Morosas"
   ClientHeight    =   6012
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7452
   HelpContextID   =   4003
   Icon            =   "frmCO_UnificacionCuotas.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6012
   ScaleWidth      =   7452
   Begin MSComCtl2.DTPicker dtpHasta 
      Height          =   315
      Left            =   5880
      TabIndex        =   20
      Top             =   960
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   550
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   194117635
      CurrentDate     =   36699
   End
   Begin MSComCtl2.DTPicker dtpDesde 
      Height          =   315
      Left            =   4080
      TabIndex        =   19
      Top             =   960
      Width           =   1215
      _ExtentX        =   2138
      _ExtentY        =   550
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   194117635
      CurrentDate     =   36699
   End
   Begin VB.Frame fraReportes 
      Height          =   3375
      Left            =   2400
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2400
         TabIndex        =   21
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   3000
         Width           =   2415
      End
      Begin MSComctlLib.ListView lswCodigos 
         Height          =   3015
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   4695
         _ExtentX        =   8276
         _ExtentY        =   5313
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   3351
         EndProperty
      End
   End
   Begin VB.ComboBox cboReportes 
      Height          =   315
      ItemData        =   "frmCO_UnificacionCuotas.frx":08CA
      Left            =   960
      List            =   "frmCO_UnificacionCuotas.frx":08D4
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   960
      Width           =   1455
   End
   Begin MSComctlLib.ProgressBar prgBarra 
      Height          =   210
      Left            =   3000
      TabIndex        =   11
      Top             =   480
      Width           =   4095
      _ExtentX        =   7218
      _ExtentY        =   360
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   720
      MaxLength       =   4
      TabIndex        =   0
      Top             =   0
      Width           =   975
   End
   Begin VB.ComboBox cboCuotas 
      Height          =   315
      ItemData        =   "frmCO_UnificacionCuotas.frx":08FD
      Left            =   720
      List            =   "frmCO_UnificacionCuotas.frx":090A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
      Width           =   735
   End
   Begin VB.TextBox txtCuotas 
      Height          =   315
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   6
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton cmdUnificar 
      Caption         =   "&Unificar"
      Height          =   315
      Left            =   6240
      TabIndex        =   5
      Top             =   5700
      Width           =   855
   End
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "&Limpiar"
      Height          =   315
      Left            =   840
      TabIndex        =   4
      Top             =   5700
      Width           =   855
   End
   Begin VB.CommandButton cmdTodas 
      Caption         =   "&Todas"
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   5700
      Width           =   855
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   360
      Width           =   855
   End
   Begin MSComctlLib.ListView lswCuotas 
      Height          =   4095
      Left            =   0
      TabIndex        =   1
      Top             =   1560
      Width           =   7095
      _ExtentX        =   12510
      _ExtentY        =   7218
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Operación"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cedula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nombre"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Saldo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Int.C"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Int.M"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Amortiza"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Estado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Ult.Mov."
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Cuotas"
         Object.Width           =   1482
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      Index           =   0
      X1              =   0
      X2              =   7080
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label lblGeneradas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Casos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   23
      Top             =   1320
      Width           =   7095
   End
   Begin VB.Label lblGeneradas 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   5280
      TabIndex        =   22
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblGeneradas 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   3480
      TabIndex        =   18
      Top             =   960
      Width           =   615
   End
   Begin VB.Label lblCodigosReporte 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Codigos >>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   2400
      TabIndex        =   14
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label lblReportes 
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   0
      TabIndex        =   12
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblCuota 
      Caption         =   "Cuotas"
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   360
      Width           =   615
   End
   Begin VB.Label lblDescripcion 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   1680
      TabIndex        =   9
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label lblCodigo 
      Caption         =   "Codigo"
      Height          =   375
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      BorderWidth     =   2
      Index           =   1
      X1              =   0
      X2              =   7080
      Y1              =   840
      Y2              =   840
   End
End
Attribute VB_Name = "frmCO_UnificacionCuotas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboCuotas_Click()
txtCuotas = "1"
End Sub

Private Sub cmdAceptar_Click()
Dim strRuta As String, strSQL As String
Dim i As Long

Me.MousePointer = vbHourglass

strSQL = ""
For i = 1 To lswCodigos.ListItems.Count
  lswCodigos.SelectedItem = lswCodigos.ListItems.Item(i)
  If lswCodigos.SelectedItem.Checked = True Then
     strSQL = strSQL & "'" & Trim(lswCodigos.SelectedItem) & "',"
  End If
Next i

If strSQL <> "" Then strSQL = Mid(strSQL, 1, Len(strSQL) - 1)


With frmContenedor.Crt
 .Reset
 .WindowShowGroupTree = True
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Reportes Módulo de Cobro"
 
 .Connect = glogon.ConectRPT
 
 If Trim(cboReportes) = "Enviados al Cobro" Then
    .Formulas(1) = "SubTitulo='CUOTAS ENVIADAS AL COBRO'"
    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_UnificacionMora.rpt")
    strSQL = "{MOROSIDAD.CODIGO} in [" & strSQL & "] and {MOROSIDAD.ESTADOI} = 'R' and"
    strSQL = strSQL & " {MOROSIDAD.FECULT} in Date(" & Year(dtpDesde) & ","
    strSQL = strSQL & Month(dtpDesde) & "," & Day(dtpDesde) & ") to Date("
    strSQL = strSQL & Year(dtpHasta) & "," & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
    .SelectionFormula = strSQL
 Else
    .Formulas(1) = "SubTitulo='ABONOS RECIBIDOS DE CUOTAS ENVIADAS AL COBRO'"
    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_UnificacionAbonos.rpt")
    strSQL = "{MOROSIDAD.CODIGO} in [" & strSQL & "] and {MOROSIDAD.ESTADOI} = 'U' and"
    strSQL = strSQL & " {MOROSIDAD.FECULT} in Date(" & Year(dtpDesde) & ","
    strSQL = strSQL & Month(dtpDesde) & "," & Day(dtpDesde) & ") to Date("
    strSQL = strSQL & Year(dtpHasta) & "," & Month(dtpHasta) & "," & Day(dtpHasta) & ")"
    .SelectionFormula = strSQL
 End If
 
 .PrintReport
End With


Me.MousePointer = vbDefault
End Sub

Private Sub cmdBuscar_Click()
Dim recMorosidad As New ADODB.Recordset
Dim strCriterio As String, strSQL As String
Dim itmX As ListItem

lswCuotas.ListItems.Clear

If Trim(cboCuotas) = "" Or Trim(txtCuotas) = "" Or Trim(txtCodigo) = "" Then
   MsgBox "Debe Suministrar el Criterio de Busqueda", vbInformation, "No Puede Continuar"
   cboCuotas.SetFocus
Else
   Me.MousePointer = vbHourglass
   
   strCriterio = cboCuotas & " " & Trim(txtCuotas)
   strSQL = "Select a.*,b.Cedula,b.Proceso,b.saldo,b.Estado,b.Fecult"
   strSQL = strSQL & " from Vista_Morosidad a,Reg_Creditos b where "
   strSQL = strSQL & " a.Codigo='" & Trim(txtCodigo) & "'"
   strSQL = strSQL & " and a.Cuota " & strCriterio
   strSQL = strSQL & " and a.Id_Solicitud=b.Id_solicitud"
   strSQL = strSQL & " and b.Proceso <> 'J'"

   With recMorosidad
     .Open strSQL, glogon.Conection, adOpenStatic
     If .RecordCount > 0 Then prgBarra.Max = .RecordCount
        Do While Not .EOF
           Set itmX = lswCuotas.ListItems.Add(, , !id_solicitud)
               itmX.SubItems(1) = Trim(!Cedula)
               itmX.SubItems(2) = fxNombre(!Cedula)
               itmX.SubItems(3) = Format(!Saldo, "Standard")
               itmX.SubItems(4) = Format(!IntC, "Standard")
               itmX.SubItems(5) = Format(!IntM, "Standard")
               itmX.SubItems(6) = Format(!Amortiza, "Standard")
               itmX.SubItems(7) = fxDescribeEstado(!Estado)
               itmX.SubItems(8) = IIf(IsNull(!FecUlt), "", Mid(!FecUlt, 1, 4) & "-" & Mid(!FecUlt, 5, 2))
               itmX.SubItems(9) = !Cuota
           .MoveNext
           prgBarra.Value = prgBarra.Value + 1
           Me.Refresh
        Loop
     prgBarra.Value = 0
     .Close
   End With
   
End If

Me.MousePointer = vbDefault

End Sub

Private Sub cmdCancelar_Click()
lswCuotas.Visible = True
fraReportes.Visible = False
txtCodigo.Enabled = True
cboCuotas.Enabled = True
txtCuotas.Enabled = True
txtCodigo.SetFocus
End Sub

Private Sub cmdLimpiar_Click()
Dim i As Long

Me.MousePointer = vbHourglass
For i = 1 To lswCuotas.ListItems.Count
    lswCuotas.ListItems.Item(i).Checked = False
Next i
Me.MousePointer = vbDefault
End Sub

Private Sub cmdTodas_Click()
Dim i As Long

Me.MousePointer = vbHourglass
For i = 1 To lswCuotas.ListItems.Count
    lswCuotas.ListItems.Item(i).Checked = True
Next i
Me.MousePointer = vbDefault
End Sub

Private Sub cmdUnificar_Click()
Dim i As Long, lngI As Long, vFecha As Date
Dim curCuotaMorosa As Currency, strSQL As String

Me.MousePointer = vbHourglass


lngI = 1
If lswCuotas.ListItems.Count > 0 Then prgBarra.Max = lswCuotas.ListItems.Count
vFecha = fxFechaServidor

For i = 1 To lswCuotas.ListItems.Count
  lswCuotas.SelectedItem = lswCuotas.ListItems(lngI)
  If lswCuotas.SelectedItem.Checked = True Then
     strSQL = "Update Morosidad set Estado='N',EstadoI='R',"
     strSQL = strSQL & "Fecult='" & Format(vFecha, "yyyy/mm/dd")
     strSQL = strSQL & "' Where Id_solicitud=" & Trim(lswCuotas.SelectedItem)
     strSQL = strSQL & " And Estado='A'"
     Call ConectionExecute(strSQL)
     
     curCuotaMorosa = CCur(lswCuotas.SelectedItem.SubItems(4)) _
                     + CCur(lswCuotas.SelectedItem.SubItems(5)) _
                     + CCur(lswCuotas.SelectedItem.SubItems(6))
     
     strSQL = "Insert into Morosidad"
     strSQL = strSQL & "(Id_solicitud,FECHAP,CUOTA_MOROSA,INTC,INTM,Amortiza,"
     strSQL = strSQL & "Estado,FECAP,ESTADOI,FECULT,Codigo)"
     strSQL = strSQL & "Values(" & Trim(lswCuotas.SelectedItem) & ","
     strSQL = strSQL & GLOBALES.glngFechaCR & "," & curCuotaMorosa & ","
     strSQL = strSQL & Format(Trim(lswCuotas.SelectedItem.SubItems(4)), "#######.00") & ","
     strSQL = strSQL & Format(Trim(lswCuotas.SelectedItem.SubItems(5)), "#######.00") & ","
     strSQL = strSQL & Format(Trim(lswCuotas.SelectedItem.SubItems(6)), "#######.00") & ","
     strSQL = strSQL & "'A'," & GLOBALES.glngFechaCR & ",'U','"
     strSQL = strSQL & Format(vFecha, "yyyy/mm/dd") & "','"
     strSQL = strSQL & UCase(Trim(txtCodigo)) & "')"
     Call ConectionExecute(strSQL)
     
     lswCuotas.ListItems.Remove (lngI)
  Else
    lngI = lngI + 1
  End If
  prgBarra.Value = prgBarra.Value + 1
  Me.Refresh
Next

lswCuotas.ListItems.Clear
prgBarra.Value = 0
Me.MousePointer = vbDefault
End Sub




Private Sub Form_Load()
dtpDesde = Format(fxFechaServidor, "dd/mm/yyyy")
dtpHasta = Format(fxFechaServidor, "dd/mm/yyyy")
Me.Width = 7230
vModulo = 4
Call Formularios(Me)
Call RefrescaTags(Me)
End Sub

Private Sub Form_Resize()

On Error Resume Next

prgBarra.Width = Me.Width - 3140
lblDescripcion.Width = Me.Width - 1800

Line1(0).X2 = Me.Width
Line1(1).X2 = Me.Width

lswCuotas.Width = Me.Width - 100
lswCuotas.Height = Me.Height - 2350
lblGeneradas(2).Width = lswCuotas.Width

cmdLimpiar.top = (lswCuotas.Height + lswCuotas.top) + 50
cmdTodas.top = (lswCuotas.Height + lswCuotas.top) + 50
cmdUnificar.top = (lswCuotas.Height + lswCuotas.top) + 50

cmdUnificar.Move lswCuotas.Left + lswCuotas.Width - (cmdUnificar.Width + 20)


End Sub

Private Sub lblCodigosReporte_Click()
Dim recCatalogo As New ADODB.Recordset
Dim itmX As ListItem, strSQL As String

If Trim(cboReportes) = "" Then
   MsgBox "Seleccione El Tipo De Reporte", vbInformation
   Exit Sub
End If

Me.MousePointer = vbHourglass

lswCuotas.ListItems.Clear
fraReportes.Visible = True
prgBarra.Visible = False
txtCodigo.Enabled = False
txtCodigo = ""
cboCuotas.Enabled = False
txtCuotas.Enabled = False
txtCuotas = ""
lblDescripcion = ""

lswCodigos.ListItems.Clear
Me.Refresh

strSQL = "Select * from Catalogo"
With recCatalogo
  .Open strSQL, glogon.Conection, adOpenStatic
    Do While .EOF = False
       Set itmX = lswCodigos.ListItems.Add(lswCodigos.ListItems.Count + 1, , Trim(!Codigo))
           itmX.SubItems(1) = Trim(!Descripcion)
           itmX.Tag = itmX.Index
       .MoveNext
    Loop
  .Close
End With

Me.MousePointer = vbDefault

End Sub


Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboCuotas.SetFocus
If KeyCode = vbKeyF4 Then
    gBusquedas.Filtro = ""
    gBusquedas.Resultado = ""
    gBusquedas.Consulta = "select codigo,descripcion from catalogo"
    gBusquedas.Orden = "codigo"
    gBusquedas.Columna = "codigo"
    frmBusquedas.Show vbModal
    txtCodigo = gBusquedas.Resultado
    If Len(Trim(txtCodigo)) > 0 Then lblDescripcion.Caption = fxDescribeCodigo(Trim(txtCodigo))
End If

End Sub

Private Sub txtCodigo_LostFocus()
If Trim(txtCodigo) <> "" Then
   lblDescripcion = fxDescribeCodigo(Trim(txtCodigo))
   If lblDescripcion = "" Then
      txtCodigo = ""
      txtCodigo.SetFocus
      lswCuotas.ListItems.Clear
   End If
End If

End Sub


Private Sub txtCuotas_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
 Case 48 To 57, 8
 Case 13
   cmdBuscar.SetFocus
 Case Else
   KeyAscii = 0
End Select
End Sub


