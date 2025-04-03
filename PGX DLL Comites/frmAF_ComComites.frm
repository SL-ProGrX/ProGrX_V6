VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmAF_ComComites 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Comites"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   7275
   Begin VB.TextBox txtCodigo 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      ToolTipText     =   "Código del Préstamo"
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtDescripcion 
      Height          =   315
      Left            =   2160
      MaxLength       =   60
      TabIndex        =   1
      ToolTipText     =   "Descripción del código de la línea"
      Top             =   480
      Width           =   4935
   End
   Begin TabDlg.SSTab ssTab 
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   5741
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Comites"
      TabPicture(0)   =   "frmAF_ComComites.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtContacto"
      Tab(0).Control(1)=   "txtFax"
      Tab(0).Control(2)=   "txtNotas"
      Tab(0).Control(3)=   "txtTelefono"
      Tab(0).Control(4)=   "cboEstado"
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(6)=   "Label1(4)"
      Tab(0).Control(7)=   "Label1(3)"
      Tab(0).Control(8)=   "Label1(2)"
      Tab(0).Control(9)=   "Label1(1)"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Actividades"
      TabPicture(1)   =   "frmAF_ComComites.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lswActividades"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Servicios"
      TabPicture(2)   =   "frmAF_ComComites.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lswServicios"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Asesores"
      TabPicture(3)   =   "frmAF_ComComites.frx":0054
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "lswAsesores"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Puestos"
      TabPicture(4)   =   "frmAF_ComComites.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "vGridPuestos"
      Tab(4).ControlCount=   1
      Begin VB.TextBox txtContacto 
         Height          =   315
         Left            =   -74040
         MaxLength       =   60
         TabIndex        =   13
         ToolTipText     =   "Código del Préstamo"
         Top             =   1080
         Width           =   5655
      End
      Begin VB.TextBox txtFax 
         Height          =   315
         Left            =   -72120
         MaxLength       =   10
         TabIndex        =   11
         ToolTipText     =   "Código del Préstamo"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtNotas 
         Height          =   1635
         Left            =   -74040
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         ToolTipText     =   "Código del Préstamo"
         Top             =   1440
         Width           =   5655
      End
      Begin VB.TextBox txtTelefono 
         Height          =   315
         Left            =   -74040
         MaxLength       =   10
         TabIndex        =   7
         ToolTipText     =   "Código del Préstamo"
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox cboEstado 
         Height          =   315
         ItemData        =   "frmAF_ComComites.frx":008C
         Left            =   -69960
         List            =   "frmAF_ComComites.frx":0096
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1575
      End
      Begin MSComctlLib.ListView lswActividades 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   15
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Actividad"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7832
         EndProperty
      End
      Begin MSComctlLib.ListView lswServicios 
         Height          =   2655
         Left            =   -74880
         TabIndex        =   16
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Servicio"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7832
         EndProperty
      End
      Begin MSComctlLib.ListView lswAsesores 
         Height          =   2655
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   4683
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Usuarios"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7832
         EndProperty
      End
      Begin FPSpreadADO.fpSpread vGridPuestos 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   18
         Top             =   360
         Width           =   6855
         _Version        =   524288
         _ExtentX        =   12091
         _ExtentY        =   4895
         _StockProps     =   64
         BackColorStyle  =   1
         BorderStyle     =   0
         EditEnterAction =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   498
         ScrollBars      =   2
         SpreadDesigner  =   "frmAF_ComComites.frx":00AC
         VScrollSpecialType=   2
         AppearanceStyle =   0
      End
      Begin VB.Label Label1 
         Caption         =   "Contacto"
         Height          =   255
         Index           =   5
         Left            =   -74760
         TabIndex        =   14
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Fax"
         Height          =   255
         Index           =   4
         Left            =   -72480
         TabIndex        =   12
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Notas"
         Height          =   255
         Index           =   3
         Left            =   -74760
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
         Height          =   255
         Index           =   2
         Left            =   -74760
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Estado"
         Height          =   255
         Index           =   1
         Left            =   -70560
         TabIndex        =   6
         Top             =   720
         Width           =   615
      End
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7275
      _ExtentX        =   12832
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "nuevo"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "editar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "borrar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "guardar"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "deshacer"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "consultar"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "reportes"
            Style           =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Código"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   510
      Width           =   615
   End
End
Attribute VB_Name = "frmAF_ComComites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vCodigo As String, vEdita As Boolean



Private Sub Form_Activate()
vModulo = 1
End Sub

Private Sub Form_Load()

On Error GoTo vError
 SSTab.Tab = 0
 vModulo = 1
 
 vEdita = True
 Call sbToolBarIconos(tlb, False)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla
 
 Call Formularios(Me)
 Call RefrescaTags(Me)
 
Exit Sub

vError:
  MsgBox Err.Description, vbExclamation
End Sub

Private Sub lswactividades_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert Afi_comactividadasg(cod_actividad,cod_comite) values('" & Item.Text _
            & "','" & vCodigo & "')"
Else
   strSQL = "Delete Afi_comactividadasg where cod_actividad ='" & Item.Text & "'"
          
End If
glogon.Conection.Execute strSQL

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub


Private Sub lswAsesores_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert Afi_comasesoresasg(cedula,cod_comite) values('" & Item.Text _
            & "','" & vCodigo & "')"
Else
   strSQL = "Delete Afi_comasesoresasg where cedula ='" & Item.Text & "'"
          
End If
glogon.Conection.Execute strSQL

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub lswServicios_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert Afi_comserviciosasg(cod_servicio,cod_comite) values('" & Item.Text _
            & "','" & vCodigo & "')"
Else
   strSQL = "Delete Afi_comserviciosasg where cod_servicio ='" & Item.Text & "'"
          
End If
glogon.Conection.Execute strSQL

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub ssTab_Click(PreviousTab As Integer)
Dim strSQL As String, rs As New ADODB.Recordset

If vCodigo = "" Then
    SSTab.Tab = 0
    Exit Sub
End If

Select Case SSTab.Tab
  Case 1 'Actividades
    Call sbCarga_Actividades
    lswActividades.SetFocus
  Case 2 'Servicios
    Call sbCargaServicios
    lswServicios.SetFocus
  Case 3 'Asesores
    Call sbCargaAsesores
    lswAsesores.SetFocus
  Case 4 'PUESTOS
    strSQL = "select cod_comite,cod_puestos,girar_desembolsos from afi_compuestosasg"
             
             '& " cod_beneficio = '" & vCodigo & "'"
    Call sbCargaGrid(vGridPuestos, 3, strSQL)
    vGridPuestos.SetFocus
End Select
End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      Call sbStabAD(False)
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
      
    Case "MODIFICAR", "EDITAR"
      SSTab.Tab = 0
      vEdita = True
      txtCodigo.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
     
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = "" Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
      
    Case "CONSULTAR"
       gBusquedas.Columna = "descripcion"
       gBusquedas.Orden = "descripcion"
       gBusquedas.Consulta = "select cod_comite,descripcion from afi_comites"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       vCodigo = gBusquedas.Resultado
       txtCodigo = gBusquedas.Resultado
       txtDescripcion = gBusquedas.Resultado2
       txtDescripcion.SetFocus
       Call sbConsulta(vCodigo)
       
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
   
End Select


End Sub

Private Sub txtCodigo_GotFocus()
txtCodigo = UCase(txtCodigo)
End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
  txtDescripcion.SetFocus
End If

If KeyCode = vbKeyF4 Then
  Call sbStabAD(False)
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_comite"
  gBusquedas.Orden = "cod_comite"
  gBusquedas.Consulta = "select cod_comite,descripcion from afi_comites"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
  txtDescripcion.SetFocus
End If

End Sub

Private Sub txtCodigo_LostFocus()
txtCodigo = UCase(txtCodigo)
End Sub

Private Sub txtContacto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtNotas.SetFocus
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyF4 Then

  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "descripcion"
  gBusquedas.Orden = "descripcion"
  gBusquedas.Consulta = "select cod_comite,descripcion from afi_comites"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  txtCodigo = gBusquedas.Resultado
  If txtCodigo <> "" Then Call sbConsulta(gBusquedas.Resultado)
End If
If KeyCode = vbKeyReturn Then txtTelefono.SetFocus
End Sub
Private Sub sbLimpiaPantalla()
vCodigo = ""
txtCodigo = ""

cboEstado.Clear
cboEstado.AddItem "ACTIVO"
cboEstado.AddItem "INACTIVO"
cboEstado.Text = "ACTIVO"

txtContacto = ""
txtNotas = ""
txtFax = ""
txtTelefono = ""
txtDescripcion = ""

Call sbStabAD(False)

End Sub



Private Sub sbConsulta(xCodigo As String)
Dim rs As New ADODB.Recordset, strSQL As String
Dim i As Integer

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from afi_comites where cod_comite = '" & xCodigo & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  vCodigo = rs!cod_comite
  txtCodigo = rs!cod_comite
 
  txtDescripcion = rs!Descripcion & ""
      
  If rs!Estado = "A" Then
    cboEstado.Text = "ACTIVO"
  Else
    cboEstado.Text = "INACTIVO"
  End If
  txtTelefono = rs!telefono
  txtFax = rs!fax
  txtContacto = rs!contacto
  txtNotas = rs!notas
  
  Call sbStabAD(True)
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close
Me.MousePointer = vbDefault
Call RefrescaTags(Me)

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub

Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

'Validar Cuentas Aqui

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripción del comite no es válido ..."


If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function

Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vEdita Then
  strSQL = "update afi_comites set descripcion = '" & UCase(Trim(txtDescripcion)) & "'" _
         & ",notas = '" & txtNotas & "',estado = '" & Mid(cboEstado.Text, 1, 1) _
         & "',telefono = '" & txtTelefono & "',fax = '" & txtFax & "',contacto = '" & txtContacto & "'" _
         & " where cod_comite = '" & vCodigo & "'"
  glogon.Conection.Execute strSQL
  Call Bitacora("Modifica", "Comite : " & vCodigo)

Else
  vCodigo = txtCodigo
   
   strSQL = "insert afi_comites(cod_comite,descripcion,estado,notas,telefono" _
          & ",fax,contacto)" _
          & " values('" & vCodigo & "','" & txtDescripcion & "','" & Mid(cboEstado.Text, 1, 1) & "','" _
          & txtNotas & "','" & txtTelefono & "','" & txtFax & "','" & txtContacto & "')"
          
   glogon.Conection.Execute strSQL
    
   Call Bitacora("Registra", "Comite: " & vCodigo)
 
End If

MsgBox "Información guardada satisfactoriamente...", vbInformation
Call sbToolBar(tlb, "activo")

Call sbStabAD(True)
Call RefrescaTags(Me)

Exit Sub

vError:
 MsgBox Err.Description, vbCritical
 
End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete afi_comites where cod_comite = '" & vCodigo & "'"
  glogon.Conection.Execute strSQL
  
  Call Bitacora("Elimina", "Comite : " & vCodigo)
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
  Call RefrescaTags(Me)
End If

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

Private Sub txtDescripcion_LostFocus()
txtDescripcion = UCase(txtDescripcion)
End Sub

Private Sub txtFax_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboEstado.SetFocus
End Sub

Private Sub txtTelefono_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then txtFax.SetFocus
End Sub



Sub sbCarga_Actividades()
Dim itmX As ListItem, i As Integer
Dim strSQL As String, rs As New ADODB.Recordset

lswActividades.ListItems.Clear

strSQL = "select A.cod_actividad,A.descripcion,C.cod_actividad as actividad" _
     & " from  Afi_comactividades A left join afi_comactividadasg C" _
     & " on A.cod_actividad = C.cod_actividad order by C.cod_actividad desc"
          
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 Set itmX = lswActividades.ListItems.Add(, , rs!Cod_actividad)
     itmX.SubItems(1) = rs!Descripcion
         
 If Not IsNull(rs!actividad) Then
     itmX.Checked = True
     itmX.ForeColor = vbBlue
 End If
 
 rs.MoveNext
Loop
rs.Close

End Sub

Private Sub sbCargaServicios()
Dim itmX As ListItem, i As Integer
Dim strSQL As String, rs As New ADODB.Recordset

lswServicios.ListItems.Clear

strSQL = "select A.cod_servicio,A.descripcion,C.cod_servicio as comite" _
     & " from  Afi_comservicios A left join afi_comserviciosasg C" _
     & " on A.cod_servicio = C.cod_servicio order by C.cod_servicio desc"
          
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 Set itmX = lswServicios.ListItems.Add(, , rs!cod_servicio)
     itmX.SubItems(1) = rs!Descripcion
         
 If Not IsNull(rs!comite) Then
     itmX.Checked = True
     itmX.ForeColor = vbBlue
 End If
 
 rs.MoveNext
Loop
rs.Close
End Sub



Private Sub sbCargaAsesores()
Dim itmX As ListItem, i As Integer
Dim strSQL As String, rs As New ADODB.Recordset

lswAsesores.ListItems.Clear

strSQL = "select A.cedula,A.nombre,C.cedula as ACedula" _
     & " from  Afi_comasesores A left join afi_comasesoresasg C" _
     & " on C.cedula = A.cedula and A.estado = 1 order by C.cedula desc"
          
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 Set itmX = lswAsesores.ListItems.Add(, , rs!Cedula)
     itmX.SubItems(1) = rs!Nombre
         
 If Not IsNull(rs!ACedula) Then
     itmX.Checked = True
     itmX.ForeColor = vbBlue
 End If
 
 rs.MoveNext
Loop
rs.Close
End Sub



Private Sub sbStabAD(vActiva As Boolean)
Dim i As Integer
'Activa o desactiva tab
 For i = 1 To SSTab.Tabs - 1
   SSTab.TabEnabled(i) = vActiva
 Next i
   
End Sub

Private Sub vGridpuestos_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Integer
'MsgBox "Columna : " & vGrid.Col
'MsgBox "Columna Activa: " & vGrid.ActiveCol
'MsgBox "Fila : " & vGrid.Row
'MsgBox "Fila Activa: " & vGrid.ActiveRow

If vGridPuestos.ActiveCol = vGridPuestos.MaxCols And (KeyCode = 13 Or KeyCode = vbKeyTab) Then
  i = fxGuardar
  If i = 0 Then Exit Sub
  vGridPuestos.Row = vGridPuestos.ActiveRow
  If vGridPuestos.MaxRows <= vGridPuestos.ActiveRow Then
    vGridPuestos.MaxRows = vGridPuestos.MaxRows + 1
    vGridPuestos.Row = vGridPuestos.MaxRows
  End If
End If

'Inserta Linea
If KeyCode = vbKeyInsert Then
    vGridPuestos.MaxRows = vGridPuestos.MaxRows + 1
    vGridPuestos.InsertRows vGridPuestos.ActiveRow, 1
    vGridPuestos.Row = vGridPuestos.ActiveRow
End If

'para buscar las opciones de puestos
If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "cod_puestos"
  gBusquedas.Orden = "cod_puestos"
  gBusquedas.Consulta = "select cod_puestos,descripcion from afi_compuestos"
  gBusquedas.Filtro = ""
  frmBusquedas.Show vbModal
  
  vGridPuestos.Row = vGridPuestos.ActiveRow
  
  vGridPuestos.Col = 1
  vGridPuestos.Text = txtCodigo
  vGridPuestos.Col = 2
  vGridPuestos.Text = gBusquedas.Resultado
  vGridPuestos.Col = 3
End If

If KeyCode = vbKeyDelete Then
   vGridPuestos.Row = vGridPuestos.ActiveRow
   vGridPuestos.Col = 2
   Call sbBorrarG(vGridPuestos.Text)
End If

End Sub


Private Function fxGuardar() As Long
Dim strSQL As String, rs As New ADODB.Recordset
'Guarda la información de la linea
'si es Insert devuelve el codigo, sino devuelve 0

On Error GoTo vError

fxGuardar = 0
vGridPuestos.Row = vGridPuestos.ActiveRow
vGridPuestos.Col = 2

strSQL = "select coalesce(count(*),0) as Existe from afi_compuestosasg " _
       & " where cod_puestos = '" & vGridPuestos.Text & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs!existe = 0 Then 'Insertar
   
  If Trim(vGridPuestos.Text) = "" Then Exit Function
  vGridPuestos.Col = 1
  strSQL = "insert into afi_compuestosasg(cod_comite,cod_puestos,girar_desembolsos) values('" _
         & UCase(vGridPuestos.Text) & "','"
  vGridPuestos.Col = 2
  strSQL = strSQL & UCase(vGridPuestos.Text) & "',"
  vGridPuestos.Col = 3
  strSQL = strSQL & vGridPuestos.Value & ")"

  glogon.Conection.Execute strSQL

  vGridPuestos.Col = 2
  Call Bitacora("Registra", "Puestos : " & vGridPuestos.Text)

Else 'Actualizar

 vGridPuestos.Col = 1
 strSQL = "update afi_compuestosasg set cod_comite = '" & vGridPuestos.Text & "',girar_desembolsos = "
 vGridPuestos.Col = 3
 strSQL = strSQL & vGridPuestos.Value & " where cod_puestos = '"
 vGridPuestos.Col = 2
 strSQL = strSQL & vGridPuestos.Text & "'"
 
 glogon.Conection.Execute strSQL

 vGridPuestos.Col = 2
 Call Bitacora("Modifica", "Puesto : " & vGridPuestos.Text)

End If
rs.Close

fxGuardar = 1

Exit Function

vError:
 MsgBox Err.Description, vbCritical

End Function


Private Sub sbBorrarG(vPuesto As String)
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  strSQL = "delete afi_compuestosasg where cod_puestos  = '" & vPuesto & "'"
  glogon.Conection.Execute strSQL
  
  Call Bitacora("Elimina", "Puesto asignado : " & vPuesto)
  Call ssTab_Click(0)
End If

Exit Sub

vError:
 MsgBox Err.Description, vbCritical

End Sub

