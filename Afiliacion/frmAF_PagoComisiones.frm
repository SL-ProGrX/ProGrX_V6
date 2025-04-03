VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmAF_PagoComisiones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago De Comisiones"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frmAF_PagoComisiones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   8250
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Reporte"
      Height          =   675
      Left            =   6840
      TabIndex        =   16
      ToolTipText     =   "Utilice la Fecha de Inicio para el Reporte"
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cboBanco 
      Height          =   315
      ItemData        =   "frmAF_PagoComisiones.frx":030A
      Left            =   960
      List            =   "frmAF_PagoComisiones.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   120
      Width           =   4695
   End
   Begin MSMask.MaskEdBox medCuenta 
      Height          =   315
      Left            =   1200
      TabIndex        =   12
      ToolTipText     =   "Cuenta Actual"
      Top             =   5040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdCuenta 
      Caption         =   "Cambiar"
      Height          =   315
      Left            =   6960
      TabIndex        =   11
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox txtNuevaCuenta 
      Height          =   315
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   10
      ToolTipText     =   "Cuenta Nueva"
      Top             =   5040
      Width           =   3735
   End
   Begin VB.CommandButton cmdAplicar 
      Caption         =   "Aplicar"
      Height          =   435
      Left            =   6960
      TabIndex        =   8
      Top             =   5640
      Width           =   1095
   End
   Begin VB.TextBox txtMonto 
      Height          =   315
      Left            =   3120
      MaxLength       =   10
      TabIndex        =   7
      Top             =   5640
      Width           =   1815
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   675
      Left            =   5880
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin MSComCtl2.DTPicker dtpInicio 
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   57147395
      CurrentDate     =   37005
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   3855
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   6800
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Promotor"
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
         SubItemIndex    =   3
         Text            =   "Afiliaciones"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "id_Banco"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Banco"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Documento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Cuenta"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Contacto"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComCtl2.DTPicker dtpCorte 
      Height          =   315
      Left            =   4320
      TabIndex        =   4
      Top             =   480
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   57147395
      CurrentDate     =   37005
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   0
      X2              =   8160
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   8160
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   8160
      X2              =   0
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C0C0C0&
      BorderWidth     =   2
      X1              =   0
      X2              =   8160
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label lblBanco 
      Caption         =   "Banco"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   855
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Index           =   1
      Left            =   0
      TabIndex        =   13
      Top             =   6120
      Width           =   8175
   End
   Begin VB.Label Label4 
      Caption         =   "Cta.Comisión"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Monto x Afiliación"
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   5640
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Corte"
      Height          =   255
      Left            =   3480
      TabIndex        =   3
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Inicio"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
End
Attribute VB_Name = "frmAF_PagoComisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbImprime(vFecha As Date)
Dim strSQL As String

With frmContenedor.Crt
  strSQL = "{SOCIOS.Fecha_Comision}= DateTime (" & Year(vFecha) & ","
  strSQL = strSQL & Month(vFecha) & "," & Day(vFecha) & ")"
  
  .Connect = glogon.ConectRPT
  
  Select Case cboBanco.ItemData(cboBanco.ListIndex)
    Case 15
      .ReportFileName = App.Path & "\Reportes\AfiPagoComisionesTE.rpt"
    Case 2
      .ReportFileName = App.Path & "\Reportes\AfiPagoComisionesCK.rpt"
    Case Else
      .ReportFileName = App.Path & "\Reportes\AfiPagoComisionesCK.rpt"
  End Select

  strSQL = strSQL & " And {BANCOS.ID_BANCO}= " & cboBanco.ItemData(cboBanco.ListIndex)
  .SelectionFormula = strSQL


  .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
  .Formulas(1) = "Monto=" & CCur(txtMonto)
  
  strSQL = "SubTitulo='Pago De Afiliaciones de " & Format(dtpInicio, "dd/mm/yyyy") _
         & " A " & Format(dtpCorte, "dd/mm/yyyy") & " Por Banco " & Trim(cboBanco.Text) & "'"
  .Formulas(2) = strSQL
  .PrintReport
End With

End Sub

Private Sub sbTesoreria(i As Integer, curMonto As Currency, intAf As Integer)
Dim recCheques As New ADODB.Recordset, rsBanco As New ADODB.Recordset
Dim lngSolicitud As Long, strCuenta As String, strSQL As String

strSQL = "Insert Into Cheques(ID_Banco,Tipo,Codigo,Beneficiario,"
strSQL = strSQL & "Monto,Fecha_Solicitud,Estado,EstadoI,Modulo,"
strSQL = strSQL & "Detalle1,Detalle2,Detalle3,CTA_AHORROS,cod_unidad,cod_concepto,user_solicita)"
strSQL = strSQL & " Values(" & lsw.ListItems.Item(i).SubItems(4) & ",'"
strSQL = strSQL & Trim(lsw.ListItems.Item(i).SubItems(6)) & "','"
strSQL = strSQL & Trim(lsw.ListItems.Item(i).SubItems(1)) & "','"
strSQL = strSQL & Mid(Trim(lsw.ListItems.Item(i).SubItems(8)), 1, 35) & "',"
strSQL = strSQL & curMonto & ",'"
strSQL = strSQL & Format(fxFechaServidor, "yyyy/mm/dd")
strSQL = strSQL & "','P','P','CC','"
strSQL = strSQL & "Pago Comision Afiliaciones" & "','"
strSQL = strSQL & "Afiliados: " & intAf & "',"
strSQL = strSQL & Mid("'Monto X Afiliacion: " & Format((curMonto / intAf), "###############.00"), 1, 27) & "','"
strSQL = strSQL & Trim(lsw.ListItems.Item(i).SubItems(7)) & "','OC','GEN','" & glogon.Usuario & "')"
glogon.Conection.Execute strSQL
    
With recCheques
 strSQL = "Select Max(NSolicitud) as Solicitud from Cheques Where Id_Banco=" & lsw.ListItems.Item(i).SubItems(4)
 strSQL = strSQL & " And Codigo='" & Trim(lsw.ListItems.Item(i).SubItems(1)) & "'"
  .Source = strSQL
  .ActiveConnection = glogon.Conection
  .CursorType = adOpenStatic
  .Open
   If .EOF = False Then
    lngSolicitud = !solicitud
   End If
 .Close
End With
     
With rsBanco
 strSQL = "Select CTACONTA from bancos where id_Banco=" & lsw.ListItems.Item(i).SubItems(4)
 .Open strSQL, glogon.Conection
   If .EOF = False Then
      strCuenta = !ctaConta
   End If
 .Close
End With
     
strSQL = "Insert Into CK_Detalle(NSolicitud,Cuenta_Contable,"
strSQL = strSQL & "Monto,DebeHaber,Linea,cod_unidad)"
strSQL = strSQL & " Values("
strSQL = strSQL & lngSolicitud & ",'" & strCuenta & "',"
strSQL = strSQL & curMonto & ",'H',"
strSQL = strSQL & 1 & ",'OC')"
glogon.Conection.Execute strSQL

strSQL = "Insert Into CK_Detalle(NSolicitud,Cuenta_Contable,"
strSQL = strSQL & "Monto,DebeHaber,Linea,cod_unidad)"
strSQL = strSQL & " Values("
strSQL = strSQL & lngSolicitud & ",'" & Trim(medCuenta) & "',"
strSQL = strSQL & curMonto & ",'D',"
strSQL = strSQL & 2 & ",'OC')"
glogon.Conection.Execute strSQL
    
Call Bitacora("Aplica", "Aplica Pago Comisiones al Promotor " & Trim(lsw.ListItems.Item(i).Text) & " " & Trim(lsw.ListItems.Item(i).SubItems(2)))
    
   
End Sub

Private Sub cmdAplicar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer
Dim curMonto As Currency
Dim intAfiliaciones As Integer
Dim vFecha As Date

Me.MousePointer = vbHourglass
vFecha = Format(fxFechaServidor, "yyyy/mm/dd")

On Error GoTo ErrorTransaccion

If Trim(txtMonto) <> "" And Trim(txtMonto) <> "." Then
   intAfiliaciones = 0
   
   For i = 1 To lsw.ListItems.Count
    lsw.SelectedItem = lsw.ListItems(i)
    If lsw.SelectedItem.Checked = True Then
       lblEstado(1) = "PROCESANDO PAGO AL PROMOTOR " & lsw.ListItems.Item(i).SubItems(2)
       Me.Refresh
       intAfiliaciones = lsw.ListItems.Item(i).SubItems(3)
       curMonto = CCur(txtMonto) * intAfiliaciones
          
       glogon.Conection.BeginTrans
            
       strSQL = "Select S.Cedula From Socios S inner join Ahorro_Consolidado A " _
              & " On S.Cedula=A.Cedula " _
              & " Where S.FechaIngreso between '" & Format(dtpInicio, "yyyy/mm/dd") _
              & "' And '" & Format(dtpCorte, "yyyy/mm/dd") & "'" _
              & " And S.Fecha_Comision is Null and S.Estadoactual <> 'N' and A.ahorro > 0 " _
              & " And id_Promotor=" & lsw.ListItems.Item(i).Text
         rs.Open strSQL, glogon.Conection, adOpenStatic
         Do While Not rs.EOF
            strSQL = "Update Socios set Fecha_Comision='" & Format(vFecha, "yyyy/mm/dd") _
                   & "' Where Cedula='" & Trim(rs!Cedula) & "'"
            glogon.Conection.Execute strSQL
            rs.MoveNext
         Loop
         rs.Close
         
       Call sbTesoreria(i, curMonto, intAfiliaciones)
       
       glogon.Conection.CommitTrans
    End If
   Next i
   
   Call sbImprime(vFecha)
Else

   MsgBox "Verifique El Monto Por Afiliación", vbExclamation
   txtMonto.SetFocus
   Me.MousePointer = vbDefault
   Exit Sub

End If

Me.MousePointer = vbDefault
lsw.ListItems.Clear
txtMonto = ""
txtNuevaCuenta = ""
lblEstado(1) = ""

Exit Sub

ErrorTransaccion:
  glogon.Conection.RollbackTrans
  Me.MousePointer = vbDefault
  lblEstado(1) = ""
  MsgBox Err.Description, vbCritical
  
End Sub

Private Sub cmdBuscar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

If Trim(cboBanco) = "" Then
   MsgBox "Especifique el banco", vbExclamation
   Exit Sub
End If

Me.MousePointer = vbHourglass
lsw.ListItems.Clear

strSQL = "Select Count(S.Cedula) as Casos,P.id_Promotor,P.Cod_Comision,P.Nombre" _
       & ",P.Cod_Banco,P.Tipo_Documento,P.Cuenta_Ahorros,P.Nombre_Contacto" _
       & " from Socios S inner join Promotores P ON P.id_promotor=S.id_promotor" _
       & " inner join Ahorro_Consolidado A on S.cedula = A.cedula" _
       & " Where S.FechaIngreso between '" & Format(dtpInicio, "yyyy/mm/dd") _
       & "' And '" & Format(dtpCorte, "yyyy/mm/dd") & "' And S.Fecha_Comision is Null" _
       & " and S.estadoactual <> 'N' and A.ahorro > 0" _
       & " And P.Cod_Banco=" & cboBanco.ItemData(cboBanco.ListIndex) _
       & " Group by P.id_promotor,P.Cod_comision,P.Nombre,P.Cod_banco,P.Tipo_Documento,P.Cuenta_Ahorros,P.Nombre_Contacto"
 
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
   Set itmX = lsw.ListItems.Add(, , rs!id_promotor)
       itmX.SubItems(1) = Trim(rs!Cod_comision)
       itmX.SubItems(2) = Trim(rs!Nombre)
       itmX.SubItems(3) = rs!casos
       itmX.SubItems(4) = rs!Cod_Banco
       itmX.SubItems(5) = fxDescribeBanco(rs!Cod_Banco)
       itmX.SubItems(6) = rs!tipo_documento
       itmX.SubItems(7) = rs!cuenta_ahorros
       itmX.SubItems(8) = rs!Nombre_Contacto
   rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub

Private Sub cmdCuenta_Click()
Dim strSQL As String

If Trim(txtNuevaCuenta) <> "" Then
 strSQL = "Update Par_afAh set Cta_Pago_Comision='" & Trim(txtNuevaCuenta) & "'"
 glogon.Conection.Execute strSQL
 medCuenta = Trim(txtNuevaCuenta)
 txtNuevaCuenta = ""
End If
End Sub

Private Sub cmdImprimir_Click()

If Trim(cboBanco) = "" Or Trim(txtMonto) = "" Then Exit Sub

Me.MousePointer = vbHourglass
Call sbImprime(dtpInicio)
Me.MousePointer = vbDefault

End Sub

Private Sub dtpCorte_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   cmdBuscar.SetFocus
End If
End Sub


Private Sub dtpInicio_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   dtpCorte.SetFocus
End If
End Sub


Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset

medCuenta.Format = GLOBALES.gstrMascara

strSQL = "Select Cta_Pago_Comision From Par_afAh"
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.EOF Then
   medCuenta = rs!Cta_Pago_Comision & ""
End If
rs.Close


strSQL = "Select id_banco,descripcion from bancos where aplica_cheques=1"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
   cboBanco.AddItem Trim(rs!Descripcion)
   cboBanco.ItemData(cboBanco.NewIndex) = rs!id_banco
   rs.MoveNext
Loop
rs.Close

dtpInicio = Format(fxFechaServidor, "dd/mm/yyyy")
dtpCorte = dtpInicio

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub

Private Sub imgBusqueda_Rapida_Click(Index As Integer)


End Sub


Private Sub txtCuenta_KeyPress(KeyAscii As Integer)

End Sub


Private Sub medCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   txtNuevaCuenta.SetFocus
Else
  KeyAscii = 0
End If
End Sub


Private Sub txtMonto_Change()
Set GLOBALES.gCajaTxt = txtMonto
Call ValidaMonto
End Sub


Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   cmdAplicar.SetFocus
End If
End Sub


Private Sub txtNuevaCuenta_KeyPress(KeyAscii As Integer)
Select Case KeyAscii
  Case 48 To 57
  Case vbKeyReturn
       cmdCuenta.SetFocus
  Case Else
       KeyAscii = 0
End Select
End Sub


