VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Begin VB.Form frmCntX_Consolidaciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Consolidaciones"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7710
   HelpContextID   =   9
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   7710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   252
      Left            =   0
      TabIndex        =   15
      Top             =   6480
      Width           =   7704
      _ExtentX        =   13600
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
   Begin TabDlg.SSTab ssTab 
      Height          =   4092
      Left            =   120
      TabIndex        =   10
      Top             =   2280
      Width           =   7332
      _ExtentX        =   12938
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Contabildades locales"
      TabPicture(0)   =   "frmCntX_Consolidaciones.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label3(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lsw"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Portales (Contabilidades Externas)"
      TabPicture(1)   =   "frmCntX_Consolidaciones.frx":0120
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "imgExp"
      Tab(1).Control(1)=   "ArbolExp"
      Tab(1).Control(2)=   "Label3(1)"
      Tab(1).ControlCount=   3
      Begin MSComctlLib.ImageList imgExp 
         Left            =   -68760
         Top             =   2280
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_Consolidaciones.frx":0248
               Key             =   "imgPortal"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_Consolidaciones.frx":0370
               Key             =   "imgRoot"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCntX_Consolidaciones.frx":0490
               Key             =   "imgConta"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lsw 
         Height          =   3255
         Left            =   120
         TabIndex        =   13
         Top             =   765
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5741
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
      Begin MSComctlLib.TreeView ArbolExp 
         Height          =   3240
         Left            =   -74880
         TabIndex        =   14
         Top             =   765
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5715
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   3
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         ImageList       =   "imgExp"
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
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seleccione las Contabilidades a Consolidar"
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
         Index           =   1
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   7095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Seleccione las Contabilidades a Consolidar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   7095
      End
   End
   Begin VB.ComboBox cboNivel 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1800
      Width           =   1575
   End
   Begin VB.TextBox txtContaBaseDesc 
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
      Locked          =   -1  'True
      TabIndex        =   6
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1440
      Width           =   5655
   End
   Begin VB.TextBox txtContaBaseCod 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   5
      ToolTipText     =   "Presione F4 para Consultar"
      Top             =   1440
      Width           =   735
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
      TabIndex        =   3
      Top             =   480
      Width           =   4935
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1200
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin MSComctlLib.Toolbar tlb 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7710
      _ExtentX        =   13600
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
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   6960
      TabIndex        =   16
      Top             =   480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   7800
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label1 
      Caption         =   "Nivel de Mascara Contable Maximo para Integración"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   3600
      TabIndex        =   9
      Top             =   1800
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Nivel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   8
      Top             =   1800
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Contabilidad Base en la que se realizara la Consolidación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   5412
   End
   Begin VB.Label Label1 
      Caption         =   "Consolidación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmCntX_Consolidaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vEdita As Boolean, vCodigo As Long, vTipoBusca As String
Dim vBusca As Boolean, vNode As Node
Dim vPaso As Boolean, vScroll As Boolean

Private Sub ArbolExp_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim strSQL As String

If vPaso Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

If Right(Node.Key, 1) = "E" Then
   If Node.Checked Then
      strSQL = "insert CNTX_CONSOLIDA_PORTALES_CON(cod_consolida,cod_portal,COD_CONTABILIDAD,Registro_Usuario,Registro_fecha) values(" _
             & vCodigo & "," & fxIndiceMixto(Node.Key, "T") & "," & fxIndiceMixto(Node.Key, "N") _
             & ",'" & glogon.Usuario & "',getdate())"
   Else
      strSQL = "delete CNTX_CONSOLIDA_PORTALES_CON where  cod_consolida = " & vCodigo _
             & " and cod_portal = " & fxIndiceMixto(Node.Key, "T") _
             & " and COD_CONTABILIDAD = " & fxIndiceMixto(Node.Key, "N")
   End If
   
   Call ConectionExecute(strSQL, 0)

End If

Me.MousePointer = vbDefault

Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If Not IsNumeric(txtCodigo.Text) Then txtCodigo.Text = 0

If vScroll Then
    strSQL = "select Top 1 COD_CONSOLIDA from CNTX_CONSOLIDA_DEFINICION"
    
    If FlatScrollBar.Value = 1 Then
       strSQL = strSQL & " where COD_CONSOLIDA > " & txtCodigo.Text & " order by COD_CONSOLIDA asc"
    Else
       strSQL = strSQL & " where COD_CONSOLIDA < " & txtCodigo.Text & " order by COD_CONSOLIDA desc"
    End If
    
    Call OpenRecordSet(rs, strSQL, 0)
    If Not rs.EOF And Not rs.BOF Then
      Call sbConsulta(rs!COD_CONSOLIDA)
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
txtContaBaseCod = ""
txtContaBaseDesc = ""

txtContaBaseDesc.Locked = True
txtCodigo.Enabled = True

txtContaBaseDesc.Enabled = True
txtContaBaseCod.Enabled = True
lsw.ListItems.Clear

cboNivel.Clear
cboNivel.AddItem "Nivel 1"
cboNivel.ItemData(cboNivel.NewIndex) = 1
cboNivel.AddItem "Nivel 2"
cboNivel.ItemData(cboNivel.NewIndex) = 2
cboNivel.AddItem "Nivel 3"
cboNivel.ItemData(cboNivel.NewIndex) = 3
cboNivel.AddItem "Nivel 4"
cboNivel.ItemData(cboNivel.NewIndex) = 4
cboNivel.AddItem "Nivel 5"
cboNivel.ItemData(cboNivel.NewIndex) = 5
cboNivel.Text = "Nivel 2"


ssTab.Tab = 0

StatusBarX.Panels(1).Text = ""
StatusBarX.Panels(2).Text = ""
StatusBarX.Panels(3).Text = ""
StatusBarX.Panels(4).Text = ""


End Sub





Private Sub ssTab_Click(PreviousTab As Integer)

If ssTab.Tab = 1 Then
   Call sbRefrescaArbol
End If

End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String

Select Case UCase(Button.Key)
    Case "INSERTAR", "NUEVO"
      vEdita = False
      Call sbLimpiaPantalla
      txtCodigo.Enabled = False
      
      txtContaBaseCod.Enabled = True
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
         gBusquedas.Columna = "cod_consolida"
         gBusquedas.Orden = "cod_consolida"
       End If
       gBusquedas.Filtro = ""
       gBusquedas.Consulta = "select cod_consolida,descripcion from CNTX_CONSOLIDA_DEFINICION"
       frmBusquedas.Show vbModal
       txtCodigo.SetFocus
       txtCodigo = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtDescripcion.SetFocus
    
    Case "REPORTES"
    
    Case "AYUDA"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
End Select

End Sub

Private Sub sbConsulta(pCodConsolida As Long)
Dim rs As New ADODB.Recordset, strSQL As String
Dim rsTmp As New ADODB.Recordset, itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "select C.cod_consolida,C.descripcion,C.nivel,E.*" _
       & ",C.REGISTRO_USUARIO,C.REGISTRO_FECHA,C.ACTUALIZA_USUARIO,C.ACTUALIZA_FECHA" _
       & " from CNTX_CONSOLIDA_DEFINICION C inner join CNTX_CONTABILIDADES E on C.COD_CONTABILIDAD = E.COD_CONTABILIDAD" _
       & " where C.cod_consolida = " & pCodConsolida
Call OpenRecordSet(rs, strSQL, 0)
If Not rs.BOF And Not rs.EOF Then
  Call sbToolBar(tlb, "activo")
  vEdita = True
  vBusca = False
  vCodigo = rs!COD_CONSOLIDA
  
  'llenar datos en pantalla
  txtCodigo = rs!COD_CONSOLIDA
  txtDescripcion = rs!Descripcion & ""
  cboNivel.Text = "Nivel " & rs!nivel
  txtContaBaseCod = rs!COD_CONTABILIDAD
  txtContaBaseDesc = rs!Nombre
  
  
    StatusBarX.Panels(1).Text = rs!Registro_Usuario & ""
    StatusBarX.Panels(2).Text = rs!REGISTRO_FECHA & ""
    StatusBarX.Panels(3).Text = rs!Actualiza_Usuario & ""
    StatusBarX.Panels(4).Text = rs!Actualiza_Fecha & ""

  vPaso = True
  
  strSQL = "select E.*" _
         & " from CNTX_CONSOLIDA_DEFINICION_DET C inner join CNTX_CONTABILIDADES E" _
         & " on C.COD_CONTABILIDAD = E.COD_CONTABILIDAD" _
         & " where C.cod_consolida = " & pCodConsolida
  rsTmp.Open strSQL, glogon.Conection, adOpenStatic
  lsw.ListItems.Clear
  Do While Not rsTmp.EOF
    txtContaBaseCod.Enabled = False
    txtContaBaseDesc.Enabled = False
    Set itmX = lsw.ListItems.Add(, , rsTmp!COD_CONTABILIDAD)
          itmX.SubItems(1) = rsTmp!Nombre
          itmX.Checked = True
    rsTmp.MoveNext
  Loop
  rsTmp.Close
  
  'Busca Otra Contabilidades con Mascara Similares no asignadas
  strSQL = "select COD_CONTABILIDAD,nombre" _
         & " from CNTX_CONTABILIDADES where nivel1 = " & rs!Nivel1 _
         & " and nivel2 = " & rs!Nivel2 & " and nivel3 = " & rs!Nivel3 _
         & " and nivel4 = " & rs!Nivel4 & " and nivel5 = " & rs!Nivel5 _
         & " and COD_CONTABILIDAD not in(select COD_CONTABILIDAD from CNTX_CONSOLIDA_DEFINICION_DET where " _
         & " cod_consolida = " & pCodConsolida & ")"
  rsTmp.Open strSQL, glogon.Conection, adOpenStatic
  Do While Not rsTmp.EOF
      Set itmX = lsw.ListItems.Add(, , rsTmp!COD_CONTABILIDAD)
          itmX.SubItems(1) = rsTmp!Nombre & ""
    rsTmp.MoveNext
  Loop
  rsTmp.Close
  
  vPaso = False
  
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

If txtDescripcion = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion de la Consolidacion no es valida ..."
If txtContaBaseDesc = "" Then vMensaje = vMensaje & vbCrLf & " - Descripcion de la Contabilidad Maestra no es valida ..."
If txtContaBaseCod = "" Then vMensaje = vMensaje & vbCrLf & " - Código de la Contabilidad Maestra no es valida ..."

If Len(vMensaje) > 0 Then
  fxValida = False
  MsgBox vMensaje, vbCritical
End If

End Function


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim lng As Long


On Error GoTo vError

If vEdita Then
  'Verificar si cambio cedula o codigo para actualización en cascada
  strSQL = "update CNTX_CONSOLIDA_DEFINICION set descripcion = '" & Trim(txtDescripcion.Text) & "',Actualiza_Usuario = '" _
         & glogon.Usuario & "',Actualiza_Fecha = getdate() where cod_consolida = " & vCodigo
  Call ConectionExecute(strSQL, 0)
  Call Bitacora("Modifica", "Consolidacion : " & vCodigo)

Else
   
   strSQL = "select isnull(max(cod_consolida),0) as ultimo from CNTX_CONSOLIDA_DEFINICION"
   Call OpenRecordSet(rs, strSQL, 0)
     txtCodigo = rs!ultimo + 1
     vCodigo = txtCodigo
   rs.Close
   
   strSQL = "insert into CNTX_CONSOLIDA_DEFINICION(descripcion,cod_consolida,COD_CONTABILIDAD,nivel,registro_usuario,registro_fecha) values('" _
          & Trim(UCase(txtDescripcion)) & "'," & vCodigo & "," & txtContaBaseCod _
          & "," & cboNivel.ItemData(cboNivel.ListIndex) & ",'" & glogon.Usuario & "',getdate())"
   Call ConectionExecute(strSQL, 0)
    
   Call Bitacora("Registra", "Consolidación: " & vCodigo)
    
   txtCodigo.Enabled = True
 
End If

'Actualizar Aqui Guarda las CONTABILIDADES Asociadas
strSQL = "delete CNTX_CONSOLIDA_DEFINICION_DET where cod_consolida = " & vCodigo
Call ConectionExecute(strSQL, 0)

For lng = 1 To lsw.ListItems.Count
  If lsw.ListItems.Item(lng).Checked Then
    strSQL = "insert into CNTX_CONSOLIDA_DEFINICION_DET(cod_consolida,COD_CONTABILIDAD,registro_usuario,registro_fecha) values(" _
           & vCodigo & "," & lsw.ListItems.Item(lng).Text & ",'" & glogon.Usuario & "',getdate())"
    Call ConectionExecute(strSQL, 0)
  End If
Next lng

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
  
  strSQL = "delete CNTX_CONSOLIDA_DEFINICION_DET where cod_consolida = " & vCodigo
  Call ConectionExecute(strSQL, 0)
  
  strSQL = "delete CNTX_CONSOLIDA_PORTALES_CON where cod_consolida = " & vCodigo
  Call ConectionExecute(strSQL, 0)

  strSQL = "delete CNTX_CONSOLIDA_HISTORIAL where cod_consolida = " & vCodigo
  Call ConectionExecute(strSQL, 0)

  strSQL = "delete CNTX_CONSOLIDA_DEFINICION where cod_consolida = " & vCodigo
  Call ConectionExecute(strSQL, 0)
  
  Call Bitacora("Elimina", "Consolidacion: " & vCodigo)

  
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
If txtCodigo <> "" And vEdita Then Call sbConsulta(txtCodigo)
End Sub

Private Sub txtContaBaseCod_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then txtContaBaseDesc.SetFocus
If KeyCode = vbKeyF4 Then
       gBusquedas.Columna = "COD_CONTABILIDAD"
       gBusquedas.Orden = "COD_CONTABILIDAD"
       gBusquedas.Filtro = ""
       gBusquedas.Consulta = "select COD_CONTABILIDAD,nombre from CNTX_CONTABILIDADES"
       frmBusquedas.Show vbModal
       txtContaBaseCod.SetFocus
       txtContaBaseCod = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtContaBaseDesc.SetFocus
End If
End Sub

Private Sub txtContaBaseCod_LostFocus()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

If Not vEdita And txtContaBaseCod <> "" Then
  lsw.ListItems.Clear
  
  strSQL = "select * from CNTX_CONTABILIDADES where COD_CONTABILIDAD = " & txtContaBaseCod
  Call OpenRecordSet(rs, strSQL, 0)
  If Not rs.EOF And Not rs.BOF Then
    txtContaBaseDesc = rs!Nombre & ""
    strSQL = "select COD_CONTABILIDAD,nombre from CNTX_CONTABILIDADES where nivel1 = " & rs!Nivel1 _
           & " and nivel2 = " & rs!Nivel2 & " and nivel3 = " & rs!Nivel3 _
           & " and nivel4 = " & rs!Nivel4 & " and nivel5 = " & rs!Nivel5
  End If
  rs.Close
  
  vPaso = True
  Call OpenRecordSet(rs, strSQL, 0)
  Do While Not rs.EOF
    Set itmX = lsw.ListItems.Add(, , rs!COD_CONTABILIDAD)
        itmX.SubItems(1) = rs!Nombre & ""
    rs.MoveNext
  Loop
  rs.Close
  vPaso = False
  
End If


Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub txtDescripcion_GotFocus()
 vTipoBusca = "D"
End Sub

Private Sub txtContaBaseDesc_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then
       gBusquedas.Columna = "nombre"
       gBusquedas.Orden = "nombre"
       gBusquedas.Filtro = ""
       gBusquedas.Consulta = "select COD_CONTABILIDAD,nombre from CNTX_CONTABILIDADES"
       frmBusquedas.Show vbModal
       txtContaBaseCod.SetFocus
       txtContaBaseCod = IIf((gBusquedas.Resultado = ""), 0, gBusquedas.Resultado)
       txtContaBaseDesc.SetFocus
End If

End Sub


Sub sbRefrescaArbol()
Dim vNode As Node, strOpciones  As String
Dim rs As New ADODB.Recordset, strSQL As String

Me.MousePointer = vbHourglass

vPaso = True

'If Not vBuscaTree Then
'  ArbolExp.Nodes.Clear
'  Exit Sub
'End If

With ArbolExp
  .Nodes.Clear
  'Crear Root
  Set vNode = .Nodes.Add(, , "Portales", "Portales", "imgRoot")
  'Crear Arbol Inicial
  
  strSQL = "select cod_portal,Descripcion from CNTX_CONSOLIDA_PORTALES"
  rs.Open strSQL, glogon.Conection, adOpenForwardOnly
  Do While Not rs.EOF
    Call sbCreaNodos("Portales", rs!Descripcion, "imgPortal", True, "0x0" & rs!cod_portal & "P")
    rs.MoveNext
  Loop
  rs.Close
  .Nodes(1).Expanded = True
End With

Me.MousePointer = vbDefault

vPaso = False

End Sub


Private Function fxIndiceCodigo(xkey As String) As String
xkey = Mid(xkey, 4, Len(xkey))
xkey = Mid(xkey, 1, Len(xkey) - 1)
fxIndiceCodigo = xkey
End Function


Private Sub ArbolExp_Expand(ByVal Node As MSComctlLib.Node)
Dim rs As New ADODB.Recordset, strSQL As String
Dim vCon As New ADODB.Connection, rsX As New ADODB.Recordset
Dim rsTmp As New ADODB.Recordset

On Error Resume Next

vPaso = True

Set vNode = Node

If Node.Tag = 1 Then Exit Sub

If Node.Index > 1 Then ArbolExp.Nodes.Remove Node.Child.Index

Node.Tag = 1

If Node.Text <> "Portales" Then

Select Case Right(Node.Key, 1)
        
    Case "P" 'Selecciona las contabilidades disponibles
             'y Que concuerdan con la mascara contable base
                 
        'Consulta la Mascara contable de la empresa Base
        strSQL = "select nivel1,nivel2,nivel3,nivel4,nivel5 from CNTX_CONTABILIDADES" _
               & " where COD_CONTABILIDAD = " & txtContaBaseCod
        rsTmp.Open strSQL, glogon.Conection, adOpenStatic
        
        strSQL = "select P.*,C.COD_CONTABILIDAD" _
               & " from CNTX_CONSOLIDA_PORTALES P inner join CNTX_CONSOLIDA_PORTALES_CONTAS C on P.cod_portal = C.cod_portal" _
               & " where P.cod_portal = " & fxIndiceCodigo(Node.Key)
        Call OpenRecordSet(rs, strSQL, 0)
        Do While Not rs.EOF
          strSQL = fxPortalPrueba(Trim(rs!por_user), fxPortalCifrado(rs!por_password, "D") _
                           , Trim(rs!por_server), Trim(rs!por_database))
          If Len(strSQL) > 0 Then
            vCon.Open strSQL
                'Selecciona las CONTABILIDADES que concuerden estrictamente con la mascara de la contabilidad base o Matriz.
                strSQL = "select COD_CONTABILIDAD,nombre from CNTX_CONTABILIDADES where COD_CONTABILIDAD = " & rs!COD_CONTABILIDAD _
                       & " and nivel1 = " & rsTmp!Nivel1 & " and nivel2 = " & rsTmp!Nivel2 _
                       & " and nivel3 = " & rsTmp!Nivel3 & " and nivel4 = " & rsTmp!Nivel4 _
                       & " and nivel5 = " & rsTmp!Nivel5
                      
                rsX.Open strSQL, vCon, adOpenStatic
                If Not rsX.EOF And Not rsX.BOF Then
                     Call sbCreaNodos(Node.Key, rsX!Nombre, "imgConta", False, "0x0" & rs!cod_portal & "-" & rs!COD_CONTABILIDAD & "E", True)
                End If
                rsX.Close
                
            vCon.Close
          End If 'Portal Abierto
          rs.MoveNext
        Loop
        
        rs.Close
        rsTmp.Close
        
End Select

End If

vPaso = False

End Sub

Private Function fxExisteContaExterna(pContabilidad As Long, pPortal As Long) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "select isnull(count(*),0) as Existe from CNTX_CONSOLIDA_PORTALES_CON" _
       & " where cod_consolida = " & vCodigo _
       & " and COD_CONTABILIDAD = " & pContabilidad _
       & " and cod_portal = " & pPortal
Call OpenRecordSet(rs, strSQL, 0)
 fxExisteContaExterna = IIf((rs!Existe > 0), True, False)
rs.Close

End Function


Private Function fxIndiceMixto(xkey As String, vTipo As String) As String
Dim i As Long, strResultado As String, blnPaso As Boolean

xkey = fxIndiceCodigo(xkey)

blnPaso = True

If vTipo = "T" Then ' Tipo
  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xkey, i, 1) <> "-" Then
     strResultado = strResultado & Mid(xkey, i, 1)
    Else
     blnPaso = False
    End If
    i = i + 1
  Loop
  
Else 'Numero

  i = 1
  strResultado = ""
  Do While blnPaso
    If Mid(xkey, i, 1) = "-" Then blnPaso = False
    i = i + 1
  Loop
  strResultado = Mid(xkey, i, 50)

End If

fxIndiceMixto = strResultado

End Function


Private Sub sbCreaNodos(vPadre As String, vTexto As String _
    , vImagen As String, vExpand As Boolean, Optional xkey As String = "N", Optional xExiste As Boolean = False)
Dim nodX As Node, vKey As String

On Error Resume Next

Set nodX = ArbolExp.Nodes.Add(vPadre, tvwChild)
    nodX.Text = vTexto
    nodX.Tag = nodX.Index
    nodX.Image = vImagen
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
    If xExiste Then
      nodX.Checked = fxExisteContaExterna(fxIndiceMixto(nodX.Key, "N"), fxIndiceMixto(nodX.Key, "T"))
    End If
    
    vKey = nodX.Key

If vExpand Then
    Set nodX = ArbolExp.Nodes.Add(vKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & ArbolExp.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If
    
End Sub


