VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFNDTrasladoASE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envio a Cobro. Traslado Sistema de Credito"
   ClientHeight    =   7464
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8172
   Icon            =   "frmFNDTrasladoASE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7464
   ScaleWidth      =   8172
   Begin VB.TextBox txtProceso 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cboInstitucion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   1080
      Width           =   6855
   End
   Begin MSComctlLib.ProgressBar prgBarra 
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   6930
      Width           =   5415
      _ExtentX        =   9546
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.CheckBox chkEnviar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enviar Todos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   1455
   End
   Begin VB.CommandButton cmdEnviar 
      Caption         =   "&Enviar al Cobro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6840
      Picture         =   "frmFNDTrasladoASE.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6360
      Width           =   1215
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   4095
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   7935
      _ExtentX        =   13991
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
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "#Contrato"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cedula"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nombre"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Plazo"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Monto"
         Object.Width           =   2469
      EndProperty
   End
   Begin VB.ComboBox cboOperadora 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1320
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
   Begin VB.TextBox txtCodigo 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtDescripcion 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   4815
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   480
      Width           =   495
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar1 
      Height          =   255
      Left            =   2640
      TabIndex        =   13
      Top             =   1440
      Width           =   495
      _ExtentX        =   868
      _ExtentY        =   445
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin VB.CommandButton cmbBuscar 
      Caption         =   "&Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5760
      Picture         =   "frmFNDTrasladoASE.frx":0486
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label lblFecha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Proceso"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Institución"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   8040
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Seleccione los contratos a enviar al cobro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   7935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Operadora"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Plan"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "frmFNDTrasladoASE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vScroll As Boolean, vPaso As Boolean
Dim vScrollPrm As Boolean


Private Sub cboInstitucion_Click()
Dim strSQL As String, rs As New ADODB.Recordset

If vPaso Then Exit Sub


strSQL = "select PR_FECHA_CORTE from INSTITUCIONES where cod_institucion = " _
        & cboInstitucion.ItemData(cboInstitucion.ListIndex)
Call OpenRecordSet(rs, strSQL)
If Not rs.EOF And Not rs.BOF Then
   txtProceso.Text = Format(rs!pr_fecha_corte, "yyyymm")
   txtCodigo_LostFocus
End If
rs.Close
End Sub

Private Sub cboOperadora_Click()

txtCodigo_LostFocus

If Trim(txtCodigo) <> "" Then Call sbConsultaContrato

End Sub

Private Sub cboOperadora_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtCodigo.SetFocus
End Sub


Private Sub chkEnviar_Click()
Dim i As Long

For i = 1 To lsw.ListItems.Count
     lsw.ListItems.Item(i).Checked = chkEnviar.Value
Next

End Sub

Private Sub cmbBuscar_Click()
 Call sbConsultaContrato
End Sub

Private Sub cmdEnviar_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Long, vCodigo As String, vOperadora As Long, vPlan As String
Dim vMonto As Currency, vFecha As String, lngOP As Long, vObservaciones As String
Dim vProceso As Currency, vProcAnterior As Currency
Dim vGarantia As String, vComite As Integer

If Trim(txtCodigo.Text) = "" Then
   MsgBox "Faltan Datos", vbExclamation
   Exit Sub
End If

On Error GoTo vError
vProceso = txtProceso.Text
vProcAnterior = fxFechaProcesoAnterior(vProceso)

vFecha = Format(fxFechaServidor, "yyyy/mm/dd")
vOperadora = cboOperadora.ItemData(cboOperadora.ListIndex)
vPlan = Trim(txtCodigo)
vCodigo = fxgFNDCodigo(vOperadora, vPlan)

vGarantia = fxCrdGarantiaOmision(vCodigo)
vComite = fxCrdIdComiteLinea(vCodigo)

With lsw.ListItems
    prgBarra.Max = .Count
    
    For i = 1 To .Count
     
     If .Item(i).Checked Then
     
         vMonto = Trim(.Item(i).SubItems(4))
         vObservaciones = "OP:" & Trim(cboOperadora.Text) & " PLN:" & Trim(txtDescripcion.Text) & " #Cont:" & Trim(.Item(i).Text)
         
         strSQL = "insert reg_creditos(codigo,id_comite,cedula,montosol,montoapr,monto_girado" _
                & ",saldo,amortiza,interesc,saldo_mes,cuota,int,interesv,plazo,userrec,userres" _
                & ",userfor,usertesoreria,tesoreria,fechasol,fechares,fechaforp,fechaforf" _
                & ",fecha_calculo_int,garantia,primer_cuota,tdocumento,ndocumento,pagare" _
                & ",firma_deudor,premio,observacion,estado,prideduc,fecult,estadosol,documento_referido)" _
                & " values('" & UCase(vCodigo) & "'," & vComite & ",'" & Trim(.Item(i).SubItems(1)) & "'," & vMonto _
                & "," & vMonto & ",0," & vMonto & ",0,0," & vMonto & "," & vMonto & ",0,0,999,'" _
                & glogon.Usuario & "','" & glogon.Usuario & "','" & glogon.Usuario & "'," & "'" & glogon.Usuario & "','" _
                & vFecha & "','" & vFecha & "','" & vFecha & "','" & vFecha & "','" & vFecha & "','" & vFecha & "','" & vGarantia & "'" _
                & ",'N','OT','',0,1,0,'" & vObservaciones & "','A'," & vProceso _
                & "," & vProcAnterior & ",'F','FND')"
           Call ConectionExecute(strSQL)
           
           lngOP = fxgFNDUltimaOperacion(Trim(.Item(i).SubItems(1)))
    
           Call sbBitacoraCredito("08", "Op: " & lngOP & " - Monto " & vMonto _
                                & " - Plazo: " & Trim(.Item(i).SubItems(3)), "R", lngOP, UCase(vCodigo))
             
           strSQL = "Update Fnd_contratos set ind_deduccion = 1, Operacion = " & lngOP _
                  & " Where cod_operadora = " & vOperadora & " and cod_plan='" & vPlan _
                  & "' and Cod_contrato=" & Trim(.Item(i).Text)
           Call ConectionExecute(strSQL)
     
     End If 'Marca
     prgBarra.Value = prgBarra.Value + 1
    
    Next i

End With


chkEnviar.Value = vbUnchecked
lsw.ListItems.Clear
prgBarra.Value = 0

MsgBox "Proceso Finalizado", vbExclamation


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
  lsw.ListItems.Clear
  prgBarra.Value = 0

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vScroll Then
    strSQL = "select Top 1 cod_plan from fnd_planes" _
           & " where cod_operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex)
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & "and deducir_planilla = 1 and cod_plan > '" & txtCodigo & "' order by cod_plan asc"
    Else
       strSQL = strSQL & "and deducir_planilla = 1 and cod_plan < '" & txtCodigo & "' order by cod_plan desc"
    End If
    
    Call OpenRecordSet(rs, strSQL)
    If Not rs.EOF And Not rs.BOF Then
      txtCodigo = rs!cod_Plan
      txtCodigo_LostFocus
    End If
    rs.Close
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBar1_Change()
Dim vFecha As Currency

On Error GoTo vError

vFecha = txtProceso.Text


If vScrollPrm Then
    
    If FlatScrollBar1.Value = 1 Then
       vFecha = fxFechaProcesoSiguiente(vFecha)
    Else
       vFecha = fxFechaProcesoAnterior(vFecha)
    End If
    
    txtProceso.Text = vFecha
      
End If



vScrollPrm = False
FlatScrollBar1.Value = 0
vScrollPrm = True


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub Form_Activate()
vModulo = 18 'Fondo de Inversion

End Sub

Private Sub Form_Load()
Dim strSQL As String

vModulo = 18 'Fondo de Inversion
 
 

Call sbgFNDCargaCombos(cboOperadora, "Operadoras")

txtProceso.Text = GLOBALES.glngFechaCR

strSQL = "select cod_institucion as IdX,rtrim(descripcion) as ItmX from instituciones order by descripcion"
Call sbLlenaCbo(cboInstitucion, strSQL, False, True)

vScroll = False
 FlatScrollBar.Value = 0
vScroll = True

vScrollPrm = False
FlatScrollBar1.Value = 0
vScrollPrm = True

Call cboInstitucion_Click

Call Formularios(Me)
Call RefrescaTags(Me)

End Sub



Private Sub sbConsultaContrato()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

Me.MousePointer = vbHourglass

lsw.ListItems.Clear

strSQL = "Select F.Cod_Contrato,F.Cedula,F.Plazo,F.Monto,S.Nombre" _
       & " From Fnd_contratos F inner join Socios S on F.cedula = S.cedula" _
       & " Where F.Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex) & " and F.ind_deduccion = 1" _
       & " and F.cod_plan = '" & Trim(txtCodigo.Text) & "' and F.Estado <> 'L' and isnull(F.operacion,0) = 0 " _
       & " and S.cod_institucion = " & cboInstitucion.ItemData(cboInstitucion.ListIndex)

Call OpenRecordSet(rs, strSQL)
Do While Not rs.EOF
  Set itmX = lsw.ListItems.Add(, , rs!COD_CONTRATO)
      itmX.SubItems(1) = Trim(rs!Cedula)
      itmX.SubItems(2) = Trim(rs!Nombre)
      itmX.SubItems(3) = rs!Plazo
      itmX.SubItems(4) = Format(rs!Monto, "Standard")
   rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

End Sub





Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "cod_plan"
   gBusquedas.Orden = "cod_plan"
   gBusquedas.Filtro = " and deducir_planilla = 1 And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   txtDescripcion.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub


Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtDescripcion.SetFocus
End Sub


Private Sub txtCodigo_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "Select Descripcion from Fnd_Planes where Cod_Operadora = " & cboOperadora.ItemData(cboOperadora.ListIndex) _
       & " And Cod_Plan='" & Trim(txtCodigo.Text) & "'"
With rs
 .Open strSQL, glogon.Conection, adOpenStatic
    If .EOF = False Then
       txtDescripcion.Text = Trim(!Descripcion)
    Else
       txtCodigo.Text = ""
       txtDescripcion.Text = ""
    End If
 .Close
End With

If Trim(txtCodigo.Text) <> "" Then Call sbConsultaContrato

End Sub


Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
   gBusquedas.Columna = "descripcion"
   gBusquedas.Orden = "descripcion"
   gBusquedas.Filtro = "and deducir_planilla = 1 And Cod_operadora=" & cboOperadora.ItemData(cboOperadora.ListIndex)
   gBusquedas.Consulta = "select cod_plan,descripcion from fnd_planes"
   frmBusquedas.Show vbModal
   lsw.SetFocus
   
   If Trim(gBusquedas.Resultado) <> "" Then
      txtCodigo = Trim(gBusquedas.Resultado)
      txtDescripcion = Trim(gBusquedas.Resultado2)
   End If
   gBusquedas.Resultado = ""
   gBusquedas.Resultado2 = ""
End If
End Sub


Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then txtCodigo_LostFocus
End Sub


Private Sub txtDescripcion_LostFocus()
If Trim(txtCodigo) <> "" Then Call sbConsultaContrato
End Sub


