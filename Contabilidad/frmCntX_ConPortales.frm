VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmCntX_ConPortales 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Portales"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab ssTab 
      Height          =   3015
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   5318
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Conexión"
      TabPicture(0)   =   "frmCntX_ConPortales.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "label1(2)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "label1(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "label1(4)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "label1(5)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3(2)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "imgProbar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtUsuario"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtClave"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtServidor"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtBaseDatos"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Contabilidades"
      TabPicture(1)   =   "frmCntX_ConPortales.frx":0124
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lsw"
      Tab(1).Control(1)=   "Label3(0)"
      Tab(1).ControlCount=   2
      Begin VB.TextBox txtBaseDatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2640
         TabIndex        =   17
         Top             =   2160
         Width           =   2535
      End
      Begin VB.TextBox txtServidor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         Left            =   2640
         TabIndex        =   16
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox txtClave 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   2640
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   1440
         Width           =   2535
      End
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2640
         TabIndex        =   14
         Top             =   1080
         Width           =   2535
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   2055
         Left            =   -74880
         TabIndex        =   8
         Top             =   765
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   3625
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Código"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descrición"
            Object.Width           =   7479
         EndProperty
      End
      Begin VB.Image imgProbar 
         Height          =   375
         Left            =   6000
         Picture         =   "frmCntX_ConPortales.frx":0220
         Stretch         =   -1  'True
         ToolTipText     =   "Probar Conexión"
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Indique los parámetros de conexión para acceder las bases de datos foráneas a Integrar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   7095
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         Caption         =   "Base de Datos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   252
         Index           =   5
         Left            =   960
         TabIndex        =   13
         Top             =   2160
         Width           =   1452
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         Caption         =   "Servidor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   4
         Left            =   960
         TabIndex        =   12
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         Caption         =   "Clave"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   11
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label label1 
         Appearance      =   0  'Flat
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   960
         TabIndex        =   10
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seleccione las Contabilidades que desea activar en el portal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   -74880
         TabIndex        =   9
         Top             =   480
         Width           =   7215
      End
   End
   Begin VB.TextBox txtNotas 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1800
      Width           =   6375
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   1440
      Width           =   5055
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   735
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   1005
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
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
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "ayuda"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   19
      Top             =   5760
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario de Registro"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3246
            MinWidth        =   3246
            Object.ToolTipText     =   "Fecha de Registro"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario que Actualiza"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3246
            MinWidth        =   3246
            Object.ToolTipText     =   "Fecha de Actualización"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   7080
      TabIndex        =   20
      Top             =   1440
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      Caption         =   "Notas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmCntX_ConPortales.frx":0482
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   7695
   End
   Begin VB.Label label1 
      Appearance      =   0  'Flat
      Caption         =   "Portal"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   1095
   End
End
Attribute VB_Name = "frmCntX_ConPortales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vTipoBusca As String
Dim vBusca As Boolean, vPaso As Boolean, vScroll As Boolean

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not IsNumeric(txtCodigo.Text) Then txtCodigo.Text = 0

If vScroll Then
    strSQL = "select Top 1 cod_portal from CNTX_CONSOLIDA_PORTALES"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where cod_portal > " & txtCodigo.Text & " order by cod_portal asc"
    Else
       strSQL = strSQL & " where cod_portal < " & txtCodigo.Text & " order by cod_portal desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!cod_portal)
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

Private Sub Form_Load()

 vEdita = True
 
 vScroll = False
 FlatScrollBar.Value = 0
 vScroll = True

 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "nuevo")
 Call sbLimpiaPantalla

 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub

Private Sub sbLimpiaPantalla()
vBusca = True
vTipoBusca = "D"
vCodigo = 0
txtCodigo = ""
txtDescripcion = ""
txtNotas = ""
txtUsuario = ""
txtClave = ""
txtServidor = ""
txtBaseDatos = ""

txtCodigo.Enabled = True

lsw.ListItems.Clear

ssTab.Tab = 0

StatusBarX.Panels(1).Text = ""
StatusBarX.Panels(2).Text = ""
StatusBarX.Panels(3).Text = ""
StatusBarX.Panels(4).Text = ""


End Sub


Private Sub imgProbar_Click()

Me.MousePointer = vbHourglass

If fxPortalPrueba(txtUsuario, txtClave, txtServidor, txtBaseDatos) = "" Then
  MsgBox "No se pudo establecer conexión...", vbExclamation
Else
  MsgBox "Conexión Satisfactoria...", vbInformation
End If

Me.MousePointer = vbDefault

End Sub

Private Sub lsw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

If Item.Checked Then
   strSQL = "insert CNTX_CONSOLIDA_PORTALES_CONTAS(cod_portal,COD_CONTABILIDAD,registro_usuario,registro_fecha) values(" _
          & vCodigo & "," & Item.Text & ",'" & glogon.Usuario & "',getdate())"
Else
   strSQL = "delete CNTX_CONSOLIDA_PORTALES_CONTAS where cod_portal = " & vCodigo _
          & " and COD_CONTABILIDAD = " & Item.Text
End If
Call ConectionExecute(strSQL, 0)

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub ssTab_Click(PreviousTab As Integer)

Select Case ssTab.Tab
 Case 1
   If vCodigo = 0 Then
     MsgBox "Debe guardar la información del portal y luego accesar a esta opción", vbInformation
     ssTab.Tab = 0
   Else
     Call sbCargaLswPortal
   End If
 Case Else
  'Nada
End Select

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = True
      
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    
    Case "MODIFICAR", "EDITAR"
      vEdita = True
      txtDescripcion.SetFocus
      Call sbToolBar(tlb, "edicion")
    Case "BORRAR"
      Call sbBorrar
    Case "GUARDAR", "SALVAR"
     If fxValida Then Call sbGuardar
    Case "DESHACER"
      Call sbToolBar(tlb, "activo")
      If vCodigo = 0 Then
        Call sbLimpiaPantalla
        Call sbToolBar(tlb, "nuevo")
        vEdita = True
      Else
        Call sbConsulta(vCodigo)
      End If
    Case "CONSULTAR"
       If vTipoBusca = "D" Then
         gBusquedas.Columna = "descripcion"
         gBusquedas.Orden = "descripcion"
       Else
         gBusquedas.Columna = "cod_portal"
         gBusquedas.Orden = "cod_portal"
       End If
       gBusquedas.Filtro = ""
       gBusquedas.Consulta = "select cod_portal,descripcion from CNTX_CONSOLIDA_PORTALES"
       frmBusquedas.Show vbModal

       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtDescripcion = IIf((gBusquedas.Resultado2 = ""), 0, gBusquedas.Resultado2)
       txtDescripcion.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
End Select

End Sub

Private Sub sbConsulta(lngCodigo As Long)
Dim rs As New ADODB.Recordset, strSQL As String

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select * from CNTX_CONSOLIDA_PORTALES where cod_portal = " & lngCodigo
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  
  ssTab.Tab = 0
  
  vBusca = False
  
  vCodigo = rs!cod_portal
  'llenar datos en pantalla
  txtCodigo = rs!cod_portal
  txtDescripcion = rs!Descripcion & ""
  txtNotas = rs!observacion
  
  txtUsuario = rs!por_user
  txtClave = fxPortalCifrado(rs!por_password, "D")
  txtServidor = rs!por_server
  txtBaseDatos = rs!por_database
  
    StatusBarX.Panels(1).Text = rs!Registro_Usuario & ""
    StatusBarX.Panels(2).Text = rs!REGISTRO_FECHA & ""
    StatusBarX.Panels(3).Text = rs!Actualiza_Usuario & ""
    StatusBarX.Panels(4).Text = rs!Actualiza_Fecha & ""
  
  
Else
  MsgBox "No se encontró registro verifique...", vbInformation
End If

rs.Close

 Call RefrescaTags(Me)

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
Private Function fxValida() As Boolean
Dim vMensaje As String

vMensaje = ""
fxValida = True

If fxPortalPrueba(txtUsuario, txtClave, txtServidor, txtBaseDatos) = "" Then vMensaje = vMensaje & vbCrLf & " - Conección del Portal no es válida ..."
If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion del Portal no es valido ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbCargaLswPortal()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCon As New ADODB.Connection, vCadena As String
Dim itmX As ListItem

Me.MousePointer = vbHourglass

On Error GoTo vError

strSQL = fxPortalPrueba(txtUsuario, txtClave, txtServidor, txtBaseDatos)

lsw.ListItems.Clear

If Len(Trim(strSQL)) = 0 Then
  MsgBox "Verificar Conección del Portal...", vbExclamation
  ssTab.Tab = 0
Else
  vCon.CommandTimeout = 300
  vCon.Open strSQL
End If

'Llena la cadena con los portales actuales
vCadena = ""
strSQL = "select COD_CONTABILIDAD from CNTX_CONSOLIDA_PORTALES_CONTAS where cod_portal = " & vCodigo
Call OpenRecordSet(rs, strSQL, 0)
Do While Not rs.EOF
 vCadena = vCadena & rs!COD_CONTABILIDAD & ","
 rs.MoveNext
Loop
rs.Close

'Si no hay contabilidades marcadas, regresar contabilidad cero
If vCadena = "" Then vCadena = "0,"

vCadena = Mid(vCadena, 1, Len(vCadena) - 1)

'Carga Lsw con las CONTABILIDADES seleccionas primero y luego todas las demas

vPaso = True

strSQL = "select COD_CONTABILIDAD,nombre from CNTX_CONTABILIDADES" _
       & " where COD_CONTABILIDAD in(" & vCadena & ")"
rs.Open strSQL, vCon, adOpenForwardOnly
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!COD_CONTABILIDAD)
     itmX.SubItems(1) = rs!Nombre
     itmX.Checked = True
 rs.MoveNext
Loop
rs.Close

strSQL = "select COD_CONTABILIDAD,nombre from CNTX_CONTABILIDADES" _
       & " where COD_CONTABILIDAD not in(" & vCadena & ")"
rs.Open strSQL, vCon, adOpenForwardOnly
Do While Not rs.EOF
 Set itmX = lsw.ListItems.Add(, , rs!COD_CONTABILIDAD)
     itmX.SubItems(1) = rs!Nombre
 rs.MoveNext
Loop
rs.Close

vPaso = False

vCon.Close

Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lng As Long


On Error GoTo vError

If vEdita Then
  strSQL = "update CNTX_CONSOLIDA_PORTALES set descripcion = '" & UCase(txtDescripcion) _
         & "',observacion = '" & txtNotas & "',por_user = '" & txtUsuario _
         & "',por_password = '" & fxPortalCifrado(txtClave, "C") _
         & "',por_server = '" & txtServidor _
         & "',por_database = '" & txtBaseDatos _
         & "',Actualiza_usuario = '" & glogon.Usuario _
         & "',Actualiza_fecha = getdate()" _
         & " where cod_portal = " & vCodigo
  Call ConectionExecute(strSQL, 0)
  
  Call Bitacora("Modifica", "Portal Codigo : " & vCodigo)

Else
  strSQL = "insert CNTX_CONSOLIDA_PORTALES(descripcion,observacion,por_user,por_password" _
         & ",por_server,por_database,registro_usuario,registro_fecha) values('" _
         & UCase(txtDescripcion) & "','" & txtNotas & "','" & txtUsuario _
         & "','" & fxPortalCifrado(txtClave, "C") & "','" & txtServidor _
         & "','" & txtBaseDatos & "','" & glogon.Usuario & "',getdate())"
  Call ConectionExecute(strSQL, 0)
  
  strSQL = "select isnull(max(cod_portal),0) as ultimo from CNTX_CONSOLIDA_PORTALES"
  Call OpenRecordSet(rs, strSQL, 0)
   vCodigo = rs!ultimo
   txtCodigo = vCodigo
  rs.Close
  
  Call Bitacora("Registra", "Portal Codigo: " & vCodigo)
  txtCodigo.Enabled = True
 
End If


MsgBox "Información guardada satisfactoriamente...", vbInformation

Call sbToolBar(tlb, "activo")

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim i As Integer, strSQL As String

On Error GoTo vError

i = MsgBox("Esta Seguro que desea borrar este registro", vbYesNo)

If i = vbYes Then
  
  strSQL = "delete CNTX_CONSOLIDA_PORTALES where cod_portal = " & vCodigo
  Call ConectionExecute(strSQL, 0)
  
  
  Call Bitacora("Elimina", "Portal Codigo : " & vCodigo)

  
  Call sbLimpiaPantalla
  Call sbToolBar(tlb, "nuevo")
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub


Private Sub txtCodigo_GotFocus()
 vTipoBusca = "C"
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then txtDescripcion.SetFocus
End Sub

Private Sub txtCodigo_LostFocus()
If txtCodigo <> "" Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtDescripcion_GotFocus()
 vTipoBusca = "D"
End Sub




