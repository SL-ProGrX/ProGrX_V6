VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmFSL_Resoluciones 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "FOSOL: Resoluciones Comite"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   15030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   15030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   9015
      Begin VB.TextBox txtExpediente 
         Alignment       =   2  'Center
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
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   3
         ToolTipText     =   "Expediente"
         Top             =   600
         Width           =   2055
      End
      Begin VB.TextBox txtCedula 
         Height          =   315
         Left            =   1320
         MaxLength       =   15
         TabIndex        =   1
         ToolTipText     =   "Presione F4 para Consultar"
         Top             =   240
         Width           =   2070
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   3375
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   5445
      End
      Begin VB.Label lblEstado 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   19
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Expediente"
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
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Identificación"
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin TabDlg.SSTab SSTab 
      Height          =   5055
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
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
      TabCaption(0)   =   "Créditos"
      TabPicture(0)   =   "frmFSL_Resoluciones.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "txtTotalSobrante"
      Tab(0).Control(1)=   "vgCreditos"
      Tab(0).Control(2)=   "Label8"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Detalle Resolución"
      TabPicture(1)   =   "frmFSL_Resoluciones.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label5"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Line1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label4(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lswComite"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "dtpFechaCausa"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtObservaciones"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "dtpFechaResolucion"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "cboResolucion"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.TextBox txtTotalSobrante 
         Alignment       =   1  'Right Justify
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
         Left            =   -62760
         Locked          =   -1  'True
         MaxLength       =   15
         TabIndex        =   20
         Top             =   4560
         Width           =   1950
      End
      Begin VB.ComboBox cboResolucion 
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
         Left            =   1335
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1260
         Width           =   4815
      End
      Begin MSComCtl2.DTPicker dtpFechaResolucion 
         Height          =   330
         Left            =   6240
         TabIndex        =   11
         Top             =   1260
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   174653443
         CurrentDate     =   41023
      End
      Begin VB.TextBox txtObservaciones 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3090
         Left            =   1320
         MultiLine       =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Observaciones"
         Top             =   1740
         Width           =   6255
      End
      Begin MSComCtl2.DTPicker dtpFechaCausa 
         Height          =   330
         Left            =   13080
         TabIndex        =   12
         Top             =   1260
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   174653443
         CurrentDate     =   41023
      End
      Begin FPSpreadADO.fpSpread vgCreditos 
         Height          =   3855
         Left            =   -74640
         TabIndex        =   14
         Top             =   600
         Width           =   14085
         _Version        =   524288
         _ExtentX        =   24844
         _ExtentY        =   6800
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
         MaxCols         =   10
         ScrollBars      =   2
         SpreadDesigner  =   "frmFSL_Resoluciones.frx":0038
         VScrollSpecialType=   2
         AppearanceStyle =   1
      End
      Begin MSComctlLib.ListView lswComite 
         Height          =   3135
         Left            =   9000
         TabIndex        =   15
         Top             =   1740
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
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
         NumItems        =   0
      End
      Begin VB.Label Label4 
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   480
         TabIndex        =   22
         Top             =   1320
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   7800
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Label Label8 
         Caption         =   "Sobrante x Aplicar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -64560
         TabIndex        =   21
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Resolución"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Cómite"
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
         Left            =   7800
         TabIndex        =   16
         Top             =   1740
         Width           =   1095
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha en la que se establece la Causa"
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
         Left            =   10080
         TabIndex        =   13
         Top             =   1260
         Width           =   3015
      End
      Begin VB.Label Label4 
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
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   1740
         Width           =   855
      End
   End
   Begin MSComctlLib.Toolbar tlbMenu 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   15030
      _ExtentX        =   26511
      _ExtentY        =   582
      ButtonWidth     =   2275
      ButtonHeight    =   582
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imgMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            Key             =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Guardar"
            Key             =   "Guardar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modif Porc"
            Key             =   "Modifica"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            Key             =   "Imprimir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Aplicar"
            Key             =   "Aplicar"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMenu 
      Left            =   14520
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Resoluciones.frx":0A5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Resoluciones.frx":72BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Resoluciones.frx":DB20
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Resoluciones.frx":14382
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Resoluciones.frx":1ABE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFSL_Resoluciones.frx":21446
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblMembresia 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   10
      Top             =   600
      Width           =   5175
   End
End
Attribute VB_Name = "frmFSL_Resoluciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strSQL As String
Dim rs As New ADODB.Recordset
Dim vFecha As String, vUsuario As String, vIdSolicitud As Long
Dim vTipoAplicacion As String, vEstadoRequisitos As String, vEstado As String
Dim vPorcentajeSug As Double, vPorcentaje As Double
Dim vSaldo As Currency, vMONTO_FOSOL As Currency
Dim vMonto As Currency
Dim vAplicado As String

Private Sub Form_Activate()
  vModulo = 22
End Sub

Private Sub Form_Load()
  vModulo = 22
  
  ssTab.Tab = 0
  ssTab.TabEnabled(1) = False
   
  vgCreditos.MaxRows = 0
  vFecha = fxFechaServidor
  vUsuario = glogon.Usuario
  dtpFechaCausa.Value = vFecha
  dtpFechaResolucion.Value = vFecha
  
  Call sbCargaResolucion
  Call sbCreaEncabezado
  
End Sub

'Crea el encabezado para la lista de miembros de comité
Private Sub sbCreaEncabezado()
 Dim vLvw As MSComctlLib.ListView
    Me.lswComite.ListItems.Clear
    Set vLvw = Me.lswComite
     vLvw.ColumnHeaders.Add , , "CODIGO", 1400
     vLvw.ColumnHeaders.Add , , "CEDULA", 1400
     vLvw.ColumnHeaders.Add , , "NOMBRE", 3000, 0
End Sub

Private Sub sbTraeResolucion(Expediente As Integer)
    strSQL = "select RESOLUCION_NOTAS from FSL_EXPEDIENTES where COD_EXPEDIENTE = '" & Expediente & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    txtObservaciones.Text = IIf(IsNull(rs!Resolucion_Notas), "", rs!Resolucion_Notas)
    rs.Close
End Sub

Private Sub sbCargaComite()
Dim vItem As MSComctlLib.ListItem
Dim vLvw As MSComctlLib.ListView
Dim vKey As String

On Error GoTo vError
   
    Me.lswComite.ListItems.Clear
    Set vLvw = Me.lswComite
    
    strSQL = "Select COD_MIEMBRO, CEDULA, NOMBRE " _
           & "from FSL_COMITE_MIEMBROS where ACTIVO = 1"
     
    rs.Open strSQL, glogon.Conection, adOpenStatic

    Do While Not rs.EOF
       vKey = Trim(rs!COD_MIEMBRO) & "(CA)"
    
       Set vItem = vLvw.ListItems.Add(, vKey, Trim(rs!COD_MIEMBRO))
                   vItem.SubItems(1) = rs!Cedula
                   vItem.SubItems(2) = rs!Nombre
       rs.MoveNext
    Loop
    rs.Close
    
Exit Sub

vError:
      MsgBox Err.Description, vbExclamation

End Sub

Private Sub sbCargaResolucion()
  cboResolucion.Clear
  cboResolucion.AddItem "RECHAZADO"
  cboResolucion.AddItem "APROBADO"
  cboResolucion.AddItem "PENDIENTE"
  cboResolucion.AddItem "APELACION"
  cboResolucion.Text = "PENDIENTE"
End Sub

Public Sub sbResolucionFOSOL(Expediente As Integer)
Dim strRuta As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

strRuta = SIFGlobal.fxSIFPathReportes("FSL_BoletaResolucion.rpt")

With frmContenedor.Crt
 .Reset
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Boleta Registro Expediente"
 .ReportFileName = strRuta
 
 .Connect = glogon.ConectRPT
 
 .SelectionFormula = "{FSL_EXPEDIENTES.COD_EXPEDIENTE}=" & Expediente
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy HH:MM:ss") & "'"
 .Formulas(1) = "Usuario='" & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"

 .Action = 1
 '.PrintReport
 
End With

Me.MousePointer = vbDefault

End Sub


Private Sub ssTab_Click(PreviousTab As Integer)
   If ssTab.Tab = 1 Then
     If Not IsNumeric(txtExpediente.Text) Then Exit Sub
     Call sbTraeResolucion(txtExpediente.Text)
     Call sbCargaComite
     Call sbTraeComiteAsignado(txtExpediente.Text)
     txtObservaciones.SetFocus
   End If
End Sub

Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Key
       
       Case "Nuevo"
          Call sbLimpiar
          txtCedula.SetFocus
          
       Case "Guardar"
         If fxValidaComiteMarcado = False Then
           MsgBox "Debe selección un miembro de Cómite que aprueba", vbCritical
           Exit Sub
         End If
         
         Call sbGuardaDetalleExpediente
         Call sbRegistroComiteEvaluador(txtExpediente.Text)
       
'       strSQL = "exec spSifRegistraTags '" & txtCedula.Text & "','" & txtExpediente.Text & "', " _
'              & " '','A04','" & glogon.Usuario & "','" & txtObservaciones.Text & "'," & txtExpediente.Text & ""
'
'       glogon.Conection.Execute strSQL
       
        MsgBox "Información Guardada Satisfactoriamente", vbInformation
        
        Call sbTraeDatosAprobacion
        
        ssTab.Tab = 0
        
       
       Case "Modifica"
         If txtCedula.Text = Empty Or txtExpediente.Text = Empty Then Exit Sub
         
         GLOBALES.gTag = txtCedula.Text
         GLOBALES.gTag2 = txtNombre.Text
         GLOBALES.gTag3 = txtExpediente.Text
         frmFSL_ModificaPorcentaje.Show vbModal
         Call sbTraeDatosAprobacion
       
'       strSQL = "exec spSifRegistraTags '" & txtCedula.Text & "','" & txtExpediente.Text & "', " _
'              & " '','A04','" & glogon.Usuario & "','" & txtObservaciones.Text & "'," & txtExpediente.Text & ""
'
'       glogon.Conection.Execute strSQL
               
       Case "Imprimir"
         Call sbResolucionFOSOL(txtExpediente.Text)
         
       Case "Aplicar"
         If txtCedula.Text = Empty Or txtExpediente.Text = Empty Or vEstado <> "APR" Or vAplicado = "S" Then Exit Sub
         GLOBALES.gTag = txtCedula.Text
         GLOBALES.gTag2 = txtNombre.Text
         GLOBALES.gTag3 = txtExpediente.Text
         frmFSL_Aplicaciones.Show vbModal
         
         Call sbTraeDatosAprobacion
      
    End Select

End Sub

Private Sub sbLimpiar()
Dim i As Integer

With vgCreditos
 .MaxRows = 1
 For i = 1 To .MaxCols
   .Col = i
   .Text = ""
 Next i
End With

txtCedula.Text = Empty
txtNombre.Text = Empty
txtExpediente.Text = Empty
lblEstado.Caption = Empty
txtObservaciones.Text = Empty
txtTotalSobrante.Text = Empty
txtExpediente.Enabled = True
 
ssTab.TabEnabled(1) = False
 
tlbMenu.Buttons.Item(2).Enabled = True
tlbMenu.Buttons.Item(3).Enabled = True

txtCedula.SetFocus

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vCedTemp As String

On Error GoTo vError

'Busca primer en el Maestro de Socios, de lo contrario revisa si es una operacion
' y regresa la cedula de la operacion
vCedTemp = ""

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
 strSQL = "select E.Cedula,S.Nombre from FSL_EXPEDIENTES E" _
        & " inner join Socios S on E.Cedula = S.Cedula" _
        & " where E.Cedula = '" & txtCedula.Text & "'"
 rs.Open strSQL, glogon.Conection, adOpenStatic
 If rs.EOF Then
   rs.Close

   Call sbLimpiar
   MsgBox "Asociado no tiene solicitud de Fondo Solidario presentada. Verifique!!!!"
   Exit Sub
 Else
   
   If vCedTemp = "" Then
     Call sbConsulta(txtCedula.Text)
   Else
     Call sbConsulta(vCedTemp)
   End If
   
   Call sbTraeExpediente
   Call sbTraeDatosAprobacion

 
 End If
 rs.Close
 

End If

If KeyCode = vbKeyF4 Then Call sbBusqueda

Exit Sub
vError:
    MsgBox Err.Description
Resume
End Sub

Private Sub sbBusqueda()

On Error GoTo vError

gBusquedas.Convertir = "N"
   
gBusquedas.Consulta = "Select E.cedula,S.nombre from fsl_Expedientes E" _
                    & " inner join Socios S on E.Cedula = S.Cedula"
gBusquedas.Columna = "nombre"
gBusquedas.Orden = "nombre"
frmBusquedas.Show vbModal
txtCedula = Trim(gBusquedas.Resultado)
gBusquedas.Consulta = ""
gBusquedas.Columna = ""
gBusquedas.Orden = ""
gBusquedas.Resultado = ""

If Trim(txtCedula) <> "" Then
    Call sbConsulta(txtCedula)
    Call sbTraeExpediente
End If

Exit Sub

vError:
  MsgBox Err.Description, vbCritical

End Sub

Private Sub txtExpediente_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  txtCedula.Text = Empty
  Call sbTraeExpediente
End If

If KeyCode = vbKeyF4 Then
    txtCedula.Text = Empty
    gBusquedas.Columna = "COD_EXPEDIENTE"
    gBusquedas.Orden = "COD_EXPEDIENTE"
    gBusquedas.Filtro = ""
    gBusquedas.Consulta = "select COD_EXPEDIENTE,CEDULA from FSL_EXPEDIENTES"
    frmBusquedas.Show vbModal
    txtExpediente = gBusquedas.Resultado
    txtCedula = gBusquedas.Resultado2
    Call sbTraeExpediente
End If

End Sub

Private Sub sbTraeExpediente()
On Error GoTo vError

If txtCedula.Text = Empty And txtExpediente.Text = Empty Then Exit Sub

strSQL = "Select COD_EXPEDIENTE, COD_PLAN, COD_CAUSA, CEDULA, DOCUMENTO_REF,NUMERO_DOC_REF, PRESENTA_CEDULA" _
       & ", PRESENTA_NOMBRE, PRESENTA_CONTACTO,MEMBRESIA_MESES, MEMBRESIA_PORCENTAJE, REQUISTOS_COMPLETOS" _
       & ", OBSERVACIONES, DETALLE_ENFERMEDAD, ENFERMEDAD_USUARIO,FECHA_ESTABLECE_CAUSA" _
       & ", REGISTRO_FECHA, REGISTRO_USUARIO, MODIFICA_USUARIO, MODIFICA_FECHA, ESTADO, RESOLUCION_ESTADO" _
       & ", RESOLUCION_NOTAS, RESOLUCION_FECHA, RESOLUCION_USUARIO, TOTAL_FOSOL, isnull(TOTAL_SOBRANTE,0) as 'TSobrante'" _
       & " from FSL_EXPEDIENTES "

If txtCedula.Text = Empty Then
  strSQL = strSQL & " where COD_EXPEDIENTE = " & txtExpediente.Text & " "
Else
  strSQL = strSQL & " where CEDULA = '" & txtCedula.Text & "' "
End If

rs.Open strSQL, glogon.Conection, adOpenStatic

If Not rs.EOF Then
  ssTab.TabEnabled(1) = True
  txtCedula.Text = rs!Cedula
  txtExpediente.Text = rs!COD_EXPEDIENTE
  vEstado = rs!Estado
  Select Case vEstado
     Case "REC"
        cboResolucion.Text = "RECHAZADO"
        lblEstado.Caption = "RECHAZADO"
        
     Case "APR"
        cboResolucion.Text = "APROBADO"
        lblEstado.Caption = "APROBADO"
        tlbMenu.Buttons.Item(2).Enabled = False
        tlbMenu.Buttons.Item(3).Enabled = False

     Case "PEN"
        cboResolucion.Text = "PENDIENTE"
        lblEstado.Caption = "PENDIENTE"
        
     Case "APL"
        cboResolucion.Text = "APELACION"
        lblEstado.Caption = "APELACION"
        
  End Select

  txtTotalSobrante = rs!TSobrante
  
  rs.Close
  
 ' Call sbConsulta(txtCedula.Text)
  Call sbTraeDatosAprobacion
  txtExpediente.Enabled = False


Else
  
  MsgBox "No Se encontró registro de la persona solicitada", vbInformation
  rs.Close
  
End If 'Not rs.EOF

Exit Sub

vError:
  MsgBox Err.Description, vbCritical
 Resume
End Sub

Private Sub sbConsulta(pCedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vFechaIng As Date, vFianzas As Boolean
Dim rsTmp As New ADODB.Recordset
Dim vEstadoSocio As String
     
vFianzas = False
    
strSQL = "select S.cedula as CedulaX,S.nombre,S.fechaingreso,S.estadoactual,S.notas,S.bloqueo,S.nota_User,S.nota_Fecha,A.*" _
       & ",dbo.fxCRDClasificacion(S.cedula,getdate()) as Clasificacion,dbo.fxSIFRatePersona(S.cedula) as Rating" _
       & ",dbo.fxSIFMensajesNumero(S.cedula) as IndMensajes, dbo.fxCBRHistorialNumero(S.cedula) as IndCobro" _
       & ",dbo.fxCBRFianzasEnMora(S.cedula) as IndFianzas" _
       & ",I.descripcion as InstitucionX, E.descripcion as EstadoX" _
       & " from socios S left join Ahorro_consolidado A on S.cedula = A.cedula" _
       & " inner join afi_estados_persona E on S.estadoActual = E.cod_estado" _
       & " inner join instituciones I on S.cod_institucion = I.cod_Institucion" _
       & " Where S.cedula = '" & Trim(pCedula) & "'"

rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic
 
If Not rs.EOF And Not rs.BOF Then
   txtCedula.Text = Trim(rs!Cedulax & "")
   txtNombre.Text = rs!Nombre & ""
     
   vFechaIng = IIf(IsNull(rs!FechaIngreso), fxFechaServidor, rs!FechaIngreso)
   lblMembresia.ForeColor = vbBlue
   lblMembresia.FontBold = False
    
   If rs!EstadoActual = "S" Then
      lblMembresia.Caption = "Membresía: " ' fxMembresia(vFechaIng)
      lblMembresia.ToolTipText = "[Ing.:" & Format(vFechaIng, "dd/mm/yyyy") & "]"
              
      'Consulta si tiene renuncia en tramite
      strSQL = "select count(*) as Existe from afi_cr_renuncias where cedula = '" & pCedula & "' and estado = 'T'"
      rsTmp.Open strSQL, glogon.Conection, adOpenStatic
      If rsTmp!existe > 0 Then
         lblMembresia.Caption = " ** Renuncia en Transito ** " & lblMembresia.Caption
         lblMembresia.ForeColor = vbRed
         lblMembresia.FontBold = True
      End If
     rsTmp.Close
    Else
     lblMembresia.Caption = "Membresía: NADA"
    End If
   
   rs.Close
       
   ssTab.Tab = 0
 End If
 
End Sub

Private Sub txtExpediente_KeyPress(KeyAscii As Integer)
  If (IsNumeric(Chr(KeyAscii)) <> True) And (KeyAscii <> 13) And (KeyAscii <> 8) Then
     KeyAscii = 0
  End If
End Sub

Private Sub sbTraeDatosAprobacion()
Dim i As Integer, vDisponible As Currency, vSobrante As Currency, vBase As Currency, vTSobrante As Currency

On Error GoTo vError

vTSobrante = 0


With vgCreditos
 .MaxRows = 1
 
    If txtExpediente.Text = Empty Then Exit Sub
       strSQL = "Select ED.COD_EXPEDIENTE, ED.PRIMERA_DEDUCCION,ED.ID_SOLICITUD " _
           & ", ED.TOTAL_DEUDA_P, ED.MONTO_LIQUIDACION, ED.PORCENTAJE, ED.MONTO_FOSOL" _
           & ", ED.TIPO_APLICACION_FOSOL, ED.MONTO_FORMALIZADO, E.ESTADO,Reg.SALDO as 'SaldoActual'" _
           & ", isnull(Vm.IntC + Vm.IntM + Vm.Cargos + Vm.Poliza,0) + Reg.Saldo as 'TDeudaActual'" _
           & ", isnull(E.Total_Sobrante,0) as 'Sobrante',E.Aplicado" _
           & " from FSL_EXPEDIENTES E inner join FSL_EXPEDIENTES_DETALLE ED on ED.COD_EXPEDIENTE = E.COD_EXPEDIENTE" _
           & "  inner join reg_creditos Reg on Ed.id_Solicitud = Reg.Id_Solicitud" _
           & "  left join vista_morosidad Vm on Reg.id_solicitud = Vm.id_Solicitud" _
           & " where E.COD_EXPEDIENTE = '" & txtExpediente.Text & "' "
    rs.Open strSQL, glogon.Conection, adOpenStatic
    
    vEstado = rs!Estado
    txtTotalSobrante.Text = Format(rs!Sobrante, "standard")
    
    vAplicado = rs!APLICADO
    
    If vAplicado = "S" Then
       lblEstado.Caption = "APLICADO"
       tlbMenu.Buttons.Item(7).Enabled = False
    Else
       Select Case rs!Estado
          Case "REC"
             lblEstado.Caption = "Rechazada"
        
          Case "APR"
             lblEstado.Caption = "Aprobada"
        
          Case "PEN"
             lblEstado.Caption = "Pendiente"
               
          Case "APL"
             lblEstado.Caption = "Apelación"
               
       End Select
    End If
    
    
    Do While Not rs.EOF
      .Row = .MaxRows
      
      If rs!Tipo_Aplicacion_Fosol = "M" Then
         vBase = rs!MONTO_FORMALIZADO
      Else
         vBase = rs!TOTAL_DEUDA_P
      End If

      
      vDisponible = vBase * rs!Porcentaje / 100
      vSobrante = rs!TDeudaActual - vDisponible
      vTSobrante = vTSobrante + vSobrante
      
      .Col = 1 'Operacion
      .Text = CStr(rs!id_Solicitud)
            
      .Col = 2 'Primer Deduccion
      .Text = Format(Format(rs!PRIMERA_DEDUCCION, "yyyymm"), "####-##")
      
      .Col = 3 'Monto
      .Text = Format(rs!MONTO_FORMALIZADO, "Standard")
      
      .Col = 4 'Total Deuda en el Momento de la Presentacion
      .Text = Format(rs!TOTAL_DEUDA_P, "Standard")
      
      .Col = 5 'Total de Deuda al momento de la resolución
      .Text = Format(rs!TDeudaActual, "Standard")
    
      .Col = 6
      .Text = Format(rs!Porcentaje, "Standard")
      
      .Col = 7 'Aplica % s/Monto
      .Value = IIf((rs!Tipo_Aplicacion_Fosol = "M"), 1, 0)
      
      .Col = 8 'Aplica % s/Total Deuda (Saldo)
      .Value = IIf((rs!Tipo_Aplicacion_Fosol <> "M"), 1, 0)
      
      .Col = 9  'Disponible
      .Text = Format(vDisponible, "Standard")
      
      .Col = 10  'Sobrante
      .Text = Format(rs!TDeudaActual, "Standard")
       
     .MaxRows = .MaxRows + 1
     rs.MoveNext
     

    Loop
   
    rs.Close
    
End With
Exit Sub
vError:
  MsgBox Err.Description, vbCritical
End Sub


Public Sub BoletaResolucion(Expediente As Integer)
Dim strRuta As String, rs As New ADODB.Recordset

Me.MousePointer = vbHourglass

strRuta = SIFGlobal.fxSIFPathReportes("FSL_BoletaResolucion.rpt")

With frmContenedor.Crt
 .Reset
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Boleta Registro Expediente"
 .ReportFileName = strRuta
 
 .Connect = glogon.ConectRPT
 
 .SelectionFormula = "{FSL_EXPEDIENTES.COD_EXPEDIENTE}=" & Expediente
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy HH:MM:ss") & "'"
 .Formulas(1) = "Usuario='" & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"

 .PrintReport
 
End With

Me.MousePointer = vbDefault

End Sub
'Tiene que estar marcado por lo menos un miembro de cómite que aprueba
Private Function fxValidaComiteMarcado() As Boolean
Dim i As Integer
    
fxValidaComiteMarcado = False

With lswComite
 For i = 1 To .ListItems.Count
    If lswComite.ListItems(i).Checked Then
       fxValidaComiteMarcado = True
    End If
 Next i
End With
    
End Function

'Guarda el detalle de la aplicacion
Private Sub sbGuardaDetalleExpediente()
Dim i As Integer, vExiste As Boolean
Dim vPorcentajeApl As Double

vExiste = True

strSQL = "UPDATE FSL_EXPEDIENTES SET FECHA_ESTABLECE_CAUSA='" & Format(dtpFechaCausa.Value, "yyyymmdd") & "' ,ESTADO = '" & Mid(cboResolucion.Text, 1, 3) & "',RESOLUCION_NOTAS = '" & txtObservaciones.Text & "', " _
       & " RESOLUCION_FECHA = '" & Format(dtpFechaResolucion.Value, "yyyymmdd") & "',RESOLUCION_USUARIO = '" & glogon.Usuario & "'" _
       & " WHERE COD_EXPEDIENTE = " & txtExpediente.Text & ""
        
glogon.Conection.Execute strSQL


With vgCreditos
For i = 1 To .MaxRows
    .Row = i
    .Col = 1
    If .Text = "" Then Exit Sub
    vIdSolicitud = IIf(.Text = Empty, 0, .Text)
    .Col = 6
    vPorcentaje = Format(.Text, "standard")
    .Col = 9
    vMONTO_FOSOL = CCur(IIf(Format(.Text, "standard") = "", 0, Format(.Text, "standard")))
    
    strSQL = "Select COD_EXPEDIENTE,ID_SOLICITUD from FSL_EXPEDIENTES_DETALLE" _
           & " where COD_EXPEDIENTE = " & txtExpediente.Text & " and ID_SOLICITUD = " & vIdSolicitud & ""
    rs.Open strSQL, glogon.Conection, adOpenStatic
         
    If rs.EOF Then
      vExiste = False
    End If
    
    rs.Close
    
    If vExiste Then
              
       If vIdSolicitud <> 0 Then
         strSQL = "Update FSL_EXPEDIENTES_DETALLE set PORCENTAJE = " & vPorcentaje & ", ACTUALIZACION_FECHA = getdate(),MODIFICA_USUARIO = '" & glogon.Usuario & "'" _
                & ", MODIFICA_FECHA= getdate(), MONTO_FOSOL = " & vMONTO_FOSOL & " where COD_EXPEDIENTE = " & txtExpediente.Text & " and ID_SOLICITUD = " & vIdSolicitud & ""
        glogon.Conection.Execute strSQL
      End If
    End If 'Existe= true
Next i

End With
End Sub

Private Sub sbRegistroComiteEvaluador(Expediente As Integer)
Dim vExpediente As Long, vCodigo As Integer
Dim i As Integer, vExiste As Boolean
On Error GoTo vError

With lswComite
  For i = 1 To .ListItems.Count
     If .ListItems(i).Checked Then
         vExpediente = Expediente
         vCodigo = .ListItems.Item(i).Text
        
        ' .SelectedItem.Text
         strSQL = " Select COD_EXPEDIENTE, COD_MIEMBRO from FSL_EXPEDIENTE_COMITE " _
                & " where COD_EXPEDIENTE = " & vExpediente & " and COD_MIEMBRO = " & vCodigo & ""
         rs.Open strSQL, glogon.Conection, adOpenStatic
         
         If rs.EOF Then
           vExiste = False
         End If
         rs.Close
         
         If vExiste = False Then
            strSQL = "Insert FSL_EXPEDIENTE_COMITE(COD_EXPEDIENTE, COD_MIEMBRO) " _
                   & " values(" & vExpediente & "," & vCodigo & ")"
            glogon.Conection.Execute strSQL
         End If
         
     Else
         strSQL = " Select COD_EXPEDIENTE, COD_MIEMBRO from FSL_EXPEDIENTE_COMITE " _
                & " where COD_EXPEDIENTE = " & vExpediente & " and COD_MIEMBRO = " & vCodigo & ""
         rs.Open strSQL, glogon.Conection, adOpenStatic
         
         If rs.EOF Then
           vExiste = False
         End If
         rs.Close
         
         If vExiste = True Then
            strSQL = "Delete FSL_EXPEDIENTE_COMITE where COD_EXPEDIENTE = " & vExpediente & " " _
                   & " and COD_MIEMBRO = " & vCodigo & ""
            glogon.Conection.Execute strSQL
         End If
         
     End If
  Next i
End With

Exit Sub

vError:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub sbTraeComiteAsignado(Expediente As Integer)
Dim vItem As MSComctlLib.ListItem
Dim vLvw As MSComctlLib.ListView
Dim vKey As String
Dim vCodigo As Integer, i As Integer


On Error GoTo vError
   
    Set vLvw = Me.lswComite
    
    With lswComite
    
    For i = 1 To .ListItems.Count
      vCodigo = .ListItems.Item(i).Text
      
      strSQL = "Select COD_MIEMBRO from FSL_EXPEDIENTE_COMITE " _
             & " where COD_EXPEDIENTE = " & txtExpediente.Text & " " _
             & " and COD_MIEMBRO = " & vCodigo & ""
      rs.Open strSQL, glogon.Conection, adOpenStatic
      
      If rs.EOF Then
        rs.Close
      Else
        If rs!COD_MIEMBRO = vCodigo Then
            .ListItems(i).Checked = True
        End If
        rs.Close
      End If
    Next i

    End With
Exit Sub

vError:
      MsgBox Err.Description, vbExclamation
End Sub
