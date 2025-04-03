VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm MDIMenu 
   BackColor       =   &H00400000&
   ClientHeight    =   6075
   ClientLeft      =   225
   ClientTop       =   825
   ClientWidth     =   9705
   Icon            =   "MDIMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageListY 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":3482
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":3D5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":4636
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":4F10
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":1A082
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":2F1F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":44366
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":594D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":6FE9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":8685C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerSalir 
      Left            =   960
      Top             =   1200
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   9705
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "tlbMain01"
      MinHeight1      =   330
      Width1          =   1470
      NewRow1         =   0   'False
      Child2          =   "tlbMain02"
      MinHeight2      =   330
      Width2          =   1005
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbMain02 
         Height          =   330
         Left            =   1665
         TabIndex        =   3
         Top             =   30
         Width           =   7950
         _ExtentX        =   14023
         _ExtentY        =   582
         ButtonWidth     =   2672
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageListY"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Reportes"
               Key             =   "Reportes"
               Object.ToolTipText     =   "Reportes Generales"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Transacción"
               Key             =   "Solicitudes"
               Object.ToolTipText     =   "Mantenimiento de Solicitudes"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Desembolsos"
               Key             =   "Desembolsos"
               Object.ToolTipText     =   "Consulta Desembolsos"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Monitoreo"
               Key             =   "Monitoreo"
               Object.ToolTipText     =   "Monitoreo de Saldos de Cuentas Bancarias"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Saldo Bancos"
               Key             =   "Saldos"
               Object.ToolTipText     =   "Configuracion del Saldo Bancario"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbMain01 
         Height          =   330
         Left            =   165
         TabIndex        =   2
         Top             =   30
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageListY"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Impresoras"
               Object.ToolTipText     =   "Configura Impresoras"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Usuario"
               Object.ToolTipText     =   "Cambio de Usuario"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bancos"
               Object.ToolTipText     =   "Listado de Bancos"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   0
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   65000
      Left            =   480
      Top             =   1200
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   5805
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   476
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Banco Actual"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            Object.Width           =   4304
            MinWidth        =   4304
            Object.ToolTipText     =   "Saldo del Banco"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   952
            MinWidth        =   952
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Servidor"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Base de Datos"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuModulo 
      Caption         =   "&Modulo"
      Begin VB.Menu mnuSeguridad 
         Caption         =   "Seguridad"
         Begin VB.Menu mnuCambiaUsuario 
            Caption         =   "Cambiar de Usuario"
         End
         Begin VB.Menu mnuCambiarContrasena 
            Caption         =   "Cambiar Contaseña"
         End
         Begin VB.Menu mnuSegSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSegCambioSO 
            Caption         =   "Cambio de S.O."
         End
      End
      Begin VB.Menu mnuSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuSolicitudes 
      Caption         =   "&Solicitudes"
      Begin VB.Menu mnuMantenimiento 
         Caption         =   "&Transacciones"
      End
      Begin VB.Menu mnuSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConsultaDesembolsos 
         Caption         =   "Consulta de Desembolsos"
      End
      Begin VB.Menu mnuReportes 
         Caption         =   "Reportes Generales"
      End
      Begin VB.Menu mnuSeparador3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUbicacion 
         Caption         =   "Traslado de Documentos"
         Begin VB.Menu mnuUbiCreaRemesas 
            Caption         =   "Creación de Remesas"
         End
         Begin VB.Menu mnuUbiRecibeRemesa 
            Caption         =   "Recepción de Remesas"
         End
         Begin VB.Menu mnuUbiSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuBeneficiarios 
            Caption         =   "Entrega Beneficiarios"
         End
         Begin VB.Menu mnuEntregasRev 
            Caption         =   "Reversar Entregas"
         End
         Begin VB.Menu mnuEntregaSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEntregasFlujo 
            Caption         =   "Reportes de Flujo"
         End
      End
      Begin VB.Menu mnuSeparador4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAutorizaciones 
         Caption         =   "Autorizaciones"
         Begin VB.Menu mnuAutorizaDocumentos 
            Caption         =   "Autorizar Documentos"
         End
         Begin VB.Menu mnuTESDesAutorizacion 
            Caption         =   "Des-Autorización de Solicitudes"
         End
         Begin VB.Menu mnuAutorizaSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuAutorizaReportes 
            Caption         =   "Reportes de Autorizaciones"
         End
      End
      Begin VB.Menu mnuBloqueos 
         Caption         =   "&Bloqueos"
      End
      Begin VB.Menu mnuProcesos 
         Caption         =   "Emisión de Documentos"
      End
      Begin VB.Menu mnuTesSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOtrasOpciones 
         Caption         =   "Otras Opciones"
         Begin VB.Menu mnuCartaControlTrans 
            Caption         =   "Carta de Control de Transferencia"
         End
         Begin VB.Menu mnuSeparadorOtras1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCopiarEsquemaSolicitud 
            Caption         =   "Copiar Esquema de Solicitud"
         End
         Begin VB.Menu mnuSeparadorOtras2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDocumentosDuplicados 
            Caption         =   "Consulta de Documentos Duplicados"
         End
         Begin VB.Menu mnuReclasificacionSolicitudes 
            Caption         =   "Reclasificación de Solicitudes"
         End
      End
      Begin VB.Menu mnuTipoEndosos 
         Caption         =   "Tipos de Endosos"
         Begin VB.Menu mnuEndosos 
            Caption         =   "Endosos"
         End
         Begin VB.Menu mnuBeneficiario 
            Caption         =   "Correccion Beneficiario"
         End
         Begin VB.Menu mnuOperacion 
            Caption         =   "Cancelación Operación"
         End
      End
      Begin VB.Menu mnuRaya 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTESMonitoreo 
         Caption         =   "Monitoreo"
         Begin VB.Menu mnuTESMonSaldos 
            Caption         =   "Monitoreo de Saldos en Bancos"
         End
         Begin VB.Menu mnuTESMonSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuTESMonConfiguracion 
            Caption         =   "Configuración del Monitoreo"
         End
      End
      Begin VB.Menu mnuTesSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenerarAsientos 
         Caption         =   "Traslado de Asientos"
      End
   End
   Begin VB.Menu mnuTesParametros 
      Caption         =   "&Parámetros"
      Begin VB.Menu mnuBancos 
         Caption         =   "Bancos"
      End
      Begin VB.Menu mnuDocumentoXBancos 
         Caption         =   "Documentos x Bancos"
      End
      Begin VB.Menu mnuParSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTiposDocumentos 
         Caption         =   "Tipos de Documentos"
      End
      Begin VB.Menu mnuParSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUnidadesNegocios 
         Caption         =   "Unidades de Negocios"
      End
      Begin VB.Menu mnuCentroCosto 
         Caption         =   "Centros de Costo"
      End
      Begin VB.Menu mnuParSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConceptosDesem 
         Caption         =   "Conceptos de Desembolsos"
      End
      Begin VB.Menu mnuUbicaciones 
         Caption         =   "Ubicaciones de Documentos"
      End
      Begin VB.Menu mnuParSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUserAutorizadores 
         Caption         =   "Usuarios Autorizadores"
      End
      Begin VB.Menu mnuCambioClaveAutoriza 
         Caption         =   "Cambio de Clave de Autorización"
      End
      Begin VB.Menu mnuParSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBancoXUsuarios 
         Caption         =   "Asignación de Accesos"
      End
      Begin VB.Menu mnuParametros 
         Caption         =   "Parámetros"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&?"
      Begin VB.Menu mnuContenido 
         Caption         =   "Contenido"
      End
      Begin VB.Menu mnuSoporteTecnico 
         Caption         =   "Soporte Técnico"
      End
      Begin VB.Menu mnuAyudaSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "Acerca De..."
      End
   End
End
Attribute VB_Name = "MDIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub MDIForm_Load()
Me.Caption = App.ProductName & " - " & App.Major & "." & App.Minor & ".r" & App.Revision

StatusBar1.Panels(3).Text = "US: " & glogon.Usuario
StatusBar1.Panels(5).Text = "SR: " & UCase(glogon.Servidor)
StatusBar1.Panels(6).Text = "DB: " & UCase(glogon.BaseDatos)

With MDIMenu.CD
   .HelpFile = App.HelpFile
   .HelpCommand = cdlHelpContext 'cdlHelpContents
End With

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then
   Cancel = True
'   Me.WindowState = 1
   TimerSalir.Interval = 50
End If
End Sub

Private Sub MDIForm_Resize()
     
If (Screen.Width \ Screen.TwipsPerPixelX) = 800 Then
   Me.Picture = frmTES_Imagenes.img800x600.Picture
Else
   Me.Picture = frmTES_Imagenes.img1024x768.Picture
End If

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
   End
End Sub

Private Sub mnuAcercaDe_Click()
 frmAcercaDe.Show vbModal
End Sub

Private Sub mnuAutorizaDocumentos_Click()
Call MuestraForms(frmTES_Autorizacion)
End Sub

Private Sub mnuAutorizaReportes_Click()
Call MuestraForms(frmTES_ReporteAutorizaciones)
End Sub

Private Sub mnuBancos_Click()
Call MuestraForms(frmTES_Bancos)
End Sub

Private Sub mnuBancoXUsuarios_Click()
Call MuestraForms(frmTES_Accesos)
End Sub

Private Sub mnuBeneficiario_Click()
Dim strSQL As String

strSQL = InputBox("Suministre el Nombre del Beneficiario", "Lease Correctamente...")


If Trim(strSQL) <> "" Then
With frmContenedor.Crt
     .Reset
     .ReportFileName = SIFGlobal.fxSIFPathReportes("TesCorreccionBeneficiario.rpt")
     .Formulas(0) = "Endoso='Lease Correctamente:  " & Trim(UCase(strSQL)) & "'"
     .Destination = crptToPrinter
     .PrintReport
End With
End If
End Sub

Private Sub mnuBeneficiarios_Click()
Call MuestraForms(frmTES_EntregaDocumentos)
End Sub

Private Sub mnuBloqueos_Click()
Call MuestraForms(frmTES_Bloqueos)
End Sub

Private Sub mnuCambiarContrasena_Click()
 frmCambiaClave.Show vbModal
End Sub

Private Sub mnuCambiaUsuario_Click()
 frmLogon.Show vbModal
    StatusBar1.Panels(3).Text = "US: " & glogon.Usuario
    StatusBar1.Panels(5).Text = "SR: " & UCase(glogon.Servidor)
    StatusBar1.Panels(6).Text = "DB: " & UCase(glogon.BaseDatos)
 Call sbXLogonOSInit(glogon.DSN)
 
End Sub

Private Sub mnuCambioClaveAutoriza_Click()
frmTES_AutorizaChKey.Show vbModal
End Sub

Private Sub mnuCartaControlTrans_Click()
 Call MuestraForms(frmTES_TransferenciaRepControl)
End Sub

Private Sub mnuCentroCosto_Click()
Call MuestraForms(frmTES_CentrosCostos)
End Sub

Private Sub mnuConceptosDesem_Click()
Call MuestraForms(frmTES_Conceptos)
End Sub

Private Sub mnuConsultaDesembolsos_Click()
Call MuestraForms(frmTES_ConsultaDesembolsos)
End Sub

Private Sub mnuContenido_Click()
   MDIMenu.CD.HelpCommand = cdlHelpContents
   MDIMenu.CD.ShowHelp
   MDIMenu.CD.HelpCommand = cdlHelpContext
End Sub

Private Sub mnuCopiarEsquemaSolicitud_Click()
Call MuestraForms(frmTES_CopiaEsquema)
End Sub

Private Sub mnuDocumentosDuplicados_Click()
Call MuestraForms(frmTES_DocumentosDup)
End Sub

Private Sub mnuDocumentoXBancos_Click()
Call MuestraForms(frmTES_BancosDoc)
End Sub

Private Sub mnuEndosos_Click()
Dim strSQL As String

strSQL = InputBox("Suministre el Endoso Del Cheque", "Endoso...")

If Trim(strSQL) <> "" Then
With frmContenedor.Crt
     .Reset
     .ReportFileName = SIFGlobal.fxSIFPathReportes("TesEndoso.rpt")
     .Formulas(0) = "Endoso='ENDOSO ESTE CHEQUE A:  " & Trim(UCase(strSQL)) & "'"
     .Destination = crptToPrinter
     .PrintReport
End With
End If
End Sub

Private Sub mnuGenerarAsientos_Click()
Call MuestraForms(frmTES_GeneraAsientos)
End Sub

Private Sub mnuMantenimiento_Click()
'' gblnTipo = False
'' frmCK_Tipo.Show vbModal
''
'' If gblnTipo = True Then
''    Unload frmCK_Tipo
''    frmCK_Solicitudes.Show vbModal
''    Unload frmCK_Solicitudes
'' End If

Call MuestraForms(frmTES_Transacciones)

End Sub

Private Sub mnuOperacion_Click()
Dim strSQL As String
Dim strRuta As String

strRuta = App.Path & "\"
strSQL = InputBox("Suministre el Numero de Operación", "Banco Popular")


If Trim(strSQL) <> "" Then
With frmContenedor.Crt
     .Reset
     .ReportFileName = strRuta + "CorreccionBeneficiario.rpt"
     .Formulas(0) = "Endoso='Cancela Operación del Banco Popular #  " & Trim(UCase(strSQL)) & "'"
     .Destination = crptToPrinter
     .PrintReport
End With
End If
End Sub

Private Sub mnuParametros_Click()
 frmTES_Parametros.Show vbModal
End Sub

Private Sub mnuProcesos_Click()

On Error GoTo vError

frmTES_EmisionDocumentos.Show vbModal

Exit Sub

vError:
  Unload frmTES_EmisionDocumentos
  frmTES_EmisionDocumentos.Show vbModal

End Sub

Private Sub mnuReclasificacionSolicitudes_Click()
frmTES_Reclasificacion.Show vbModal
End Sub


Private Sub mnuReportes_Click()
Call MuestraForms(frmTES_Reportes)
End Sub

Private Sub mnuSalir_Click()
 End
End Sub

Private Sub mnuSegCambioSO_Click()
frmLogonOS.Show vbModal
Call sbXLogonOSInit(glogon.DSN)
End Sub

Private Sub mnuTESDesAutorizacion_Click()
Call MuestraForms(frmTES_DesAutorizaciones)
End Sub

Private Sub mnuTESMonConfiguracion_Click()
Call MuestraForms(frmTES_BancosSaldos)
End Sub

Private Sub mnuTESMonSaldos_Click()
frmTES_Monitoreo.Show
End Sub

Private Sub mnuTiposDocumentos_Click()
Call MuestraForms(frmTES_Documentos)
End Sub

Private Sub mnuUbicaciones_Click()
Call MuestraForms(frmTES_Ubicaciones)
End Sub

Private Sub mnuUbiCreaRemesas_Click()
Call MuestraForms(frmTES_TrasladosDocumentos)
End Sub

Private Sub mnuUbiRecibeRemesa_Click()
Call MuestraForms(frmTES_RecepcionDocumentos)
End Sub

Private Sub mnuUnidadesNegocios_Click()
Call MuestraForms(frmTES_Unidades)
End Sub

Private Sub mnuUserAutorizadores_Click()
Call MuestraForms(frmTES_Autorizadores)
End Sub

Private Sub TimerSalir_Timer()
TimerSalir.Interval = 0
Call mnuSalir_Click
End Sub

Private Sub tlbMain01_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
 Case "Explorer"
'   frmCK_Explorer.Show
   MsgBox "Explorador Desactivado hasta nuevo aviso...", vbInformation
 Case "Impresoras"
   Call MuestraForms(frmTES_Impresoras)
 Case "Usuario"
  Call mnuCambiaUsuario_Click
 Case "Bancos"
  Call MuestraForms(frmTES_BancoLista)
End Select
End Sub

Private Sub tlbMain02_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
 Case "Reportes"
  Call mnuReportes_Click
 Case "Solicitudes"
  Call mnuMantenimiento_Click
 Case "Monitoreo"
   frmTES_Monitoreo.Show
 Case "Desembolsos"
   Call MuestraForms(frmTES_ConsultaDesembolsos)
 Case "Saldos"
   Call MuestraForms(frmTES_BancosSaldos)
End Select
End Sub
