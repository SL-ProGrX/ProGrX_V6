VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.MDIForm MDIMenu 
   BackColor       =   &H00808080&
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   13260
   Icon            =   "MDIMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerSalir 
      Left            =   600
      Top             =   480
   End
   Begin VB.Timer tmr1 
      Interval        =   10
      Left            =   120
      Top             =   480
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":3482
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":9CE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":10546
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":16DA8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":1D60A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":23E6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":2A6CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":2A7F6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar clb 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   688
      EmbossPicture   =   -1  'True
      _CBWidth        =   13260
      _CBHeight       =   390
      _Version        =   "6.7.9782"
      Child1          =   "tlbPeriodo"
      MinHeight1      =   330
      Width1          =   4995
      NewRow1         =   0   'False
      Child2          =   "tlbCierre"
      MinHeight2      =   330
      Width2          =   1815
      NewRow2         =   0   'False
      BandBackColor3  =   -2147483646
      Child3          =   "tlbEmpresa"
      MinWidth3       =   4005
      MinHeight3      =   330
      Width3          =   4005
      NewRow3         =   0   'False
      Begin VB.TextBox lblPeriodo 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Height          =   215
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   100
         Width           =   2535
      End
      Begin MSComctlLib.Toolbar tlbEmpresa 
         Height          =   330
         Left            =   7035
         TabIndex        =   6
         Top             =   30
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   582
         ButtonWidth     =   1138
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Empresa"
               Object.ToolTipText     =   "Contabilidad Actual"
               ImageIndex      =   6
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbCierre 
         Height          =   330
         Left            =   5190
         TabIndex        =   5
         Top             =   30
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   582
         ButtonWidth     =   1667
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cierres"
               Key             =   "cierres"
               ImageIndex      =   1
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   3
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "CierrePeriodo"
                     Text            =   "Periodo"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "CerSep1"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "CierreFiscal"
                     Text            =   "Asientos Fiscal"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtAnio 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   3
         Top             =   80
         Width           =   615
      End
      Begin VB.TextBox txtMes 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   80
         Width           =   375
      End
      Begin MSComctlLib.Toolbar tlbPeriodo 
         Height          =   330
         Left            =   165
         TabIndex        =   4
         Top             =   30
         Width           =   4800
         _ExtentX        =   8467
         _ExtentY        =   582
         ButtonWidth     =   1931
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Periodos"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2940
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "02:22 p.m."
            Object.ToolTipText     =   "Hora"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Usuario Activo"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "04/03/2016"
            Object.ToolTipText     =   "Fecha del Sistema"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   989
            MinWidth        =   989
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Tecla de CAP activa"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   988
            MinWidth        =   988
            TextSave        =   "NUM"
            Object.ToolTipText     =   "NumLock"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9596
            MinWidth        =   9596
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
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuExplorador 
         Caption         =   "Explorador"
      End
      Begin VB.Menu mnuUtilitarios 
         Caption         =   "Utilitarios"
         Begin VB.Menu mnuVerificaAsientos 
            Caption         =   "Verificación de Asientos"
         End
         Begin VB.Menu mnuEliminaAsientos 
            Caption         =   "Eliminación de Asientos"
         End
         Begin VB.Menu mnuEliminaContabilidad 
            Caption         =   "Eliminación de Contabilidades"
         End
         Begin VB.Menu mnuUtilitariosSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuRestructurarMovimientosCuentas 
            Caption         =   "Restructurar Movimientos por Cuenta"
         End
         Begin VB.Menu mnuCopiarEsquema 
            Caption         =   "Copiar Esquema a Otra Contabilidad"
         End
         Begin VB.Menu mnuUtilitariosSep2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuProcesosAdd 
            Caption         =   "Procesos Adicionales"
         End
      End
      Begin VB.Menu mnuArchivoSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSeguridad 
         Caption         =   "&Seguridad"
         Begin VB.Menu mnuCambiaClave 
            Caption         =   "Cambiar Clave"
         End
         Begin VB.Menu mnuCambioUsuario 
            Caption         =   "Cambio de Usuario"
         End
         Begin VB.Menu mnuSeguridadSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCambioSO 
            Caption         =   "Cambio de S.O."
         End
      End
      Begin VB.Menu mnuArchivoSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuExplorerContable 
      Caption         =   "Explorar"
      Visible         =   0   'False
      Begin VB.Menu mnuCntAccionEditar 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnuCntAccionBorrar 
         Caption         =   "Borrar"
      End
      Begin VB.Menu mnuAccionSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCntAccionRefrescar 
         Caption         =   "Refrescar "
      End
      Begin VB.Menu mnuAccionSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCntAccionesImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuAccionSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCntAccionesMayorizar 
         Caption         =   "Mayorizar"
      End
   End
   Begin VB.Menu mnuContabilidad 
      Caption         =   "&Contabilidad"
      Begin VB.Menu mnuCatalogoCuentas 
         Caption         =   "Catálogo de Cuentas"
      End
      Begin VB.Menu mnuContaSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRegistroAsientos 
         Caption         =   "Registro de Asientos"
      End
      Begin VB.Menu mnuAutorizacionAsientos 
         Caption         =   "Autorización de Asientos"
      End
      Begin VB.Menu mnuContaSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlantillaS 
         Caption         =   "&Plantillas"
         Begin VB.Menu mnuPlantillasDefinicion 
            Caption         =   "Asientos Fijos y Proyecciones (Definición)"
         End
         Begin VB.Menu mnuPlantillasGeneracion 
            Caption         =   "Asientos Fijos y Proyecciones (Generación)"
         End
         Begin VB.Menu mnuPlantillaSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPlantillaAsientoRate 
            Caption         =   "Asientos Porcentuales (Definición)"
         End
         Begin VB.Menu mnuPlantillaAsientosRateGen 
            Caption         =   "Asientos Porcentuales (Generación)"
         End
      End
      Begin VB.Menu mnuAreasDeTrabajo 
         Caption         =   "Areas de &Trabajo"
         Begin VB.Menu mnuAreasTrabajoDefinicion 
            Caption         =   "Definición"
         End
         Begin VB.Menu mnuAreasTrabajoReportes 
            Caption         =   "Reportes"
         End
      End
      Begin VB.Menu mnuAdmDiferidos 
         Caption         =   "Administración de Diferidos"
         Begin VB.Menu mnuDiferidosPlanillas 
            Caption         =   "Plantillas"
         End
         Begin VB.Menu mnuDiferidosSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuDiferidosCreacion 
            Caption         =   "Creación de Mov. Diferidos"
         End
         Begin VB.Menu mnuDiferidosGeneracion 
            Caption         =   "Generación de Mov. Diferidos"
         End
      End
      Begin VB.Menu mnuContaSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRazonesFinancieras 
         Caption         =   "Razones Financieras"
      End
      Begin VB.Menu mnuContaSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTipoDeCambio 
         Caption         =   "Tipo de Cambio"
      End
      Begin VB.Menu mnuDiferencialCambiario 
         Caption         =   "Diferencial Cambiario"
      End
      Begin VB.Menu mnuContaSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMayorización 
         Caption         =   "Mayorización General"
      End
      Begin VB.Menu mnuCambioEmpresa 
         Caption         =   "Cambio de Contabilidad"
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "&Reportes"
      Begin VB.Menu mnuRepMovTipoDocumento 
         Caption         =   "Documentos"
      End
      Begin VB.Menu mnuRepMovTipoCuenta 
         Caption         =   "Analítico de Cuentas"
      End
      Begin VB.Menu mnuReportesSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRepBalanceComprobacionPosCierre 
         Caption         =   "Balance de Comprobacion "
      End
      Begin VB.Menu mnuRepBalanceSit 
         Caption         =   "Balance de Situación"
      End
      Begin VB.Menu mnuReportesSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCntX_ReporteBalanceGenRes 
         Caption         =   "Balance General - Estado de Resultados"
      End
      Begin VB.Menu mnuCntX_ReporteBalanceRsm 
         Caption         =   "Resumen del Balance"
      End
      Begin VB.Menu mnuERSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportePersonalizado 
         Caption         =   "Reportes Personalizados"
      End
      Begin VB.Menu mnuRepMovPeriodo 
         Caption         =   "Movimientos del Periodo Ordinario"
      End
   End
   Begin VB.Menu mnuProfesional 
      Caption         =   "&Profesional"
      Begin VB.Menu mnuProConsolidaciones 
         Caption         =   "Consolidación: Definición"
      End
      Begin VB.Menu mnuProConsolidaPortales 
         Caption         =   "Portales (Enlace Bases de Datos Externas)"
      End
      Begin VB.Menu mnuProConCierre 
         Caption         =   "Consolidación: Integración"
      End
      Begin VB.Menu mnuProSep1 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuProPresupuesto 
         Caption         =   "Presupuesto"
         Begin VB.Menu mnuPresupuestoCuentas 
            Caption         =   "Asignación Presupuestario x Cuentas"
         End
         Begin VB.Menu mnuPresupuestoReportes 
            Caption         =   "Reportes Presupuestarios"
         End
         Begin VB.Menu mnuPreupuestoSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPresupuestoLotes 
            Caption         =   "Asignación Presupuestaria x Lotes"
         End
      End
      Begin VB.Menu mnuProPresuControl 
         Caption         =   "Control del Presupuesto"
      End
      Begin VB.Menu mnuProAsientosPre 
         Caption         =   "Asientos Presupuestarios"
         Begin VB.Menu mnuProPreAsientos 
            Caption         =   "Asientos (Confección)"
         End
         Begin VB.Menu mnuProPreAsientoSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuProPreAsientosRep 
            Caption         =   "Reportes de Asientos Pre."
         End
      End
      Begin VB.Menu mnuProSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuProRastreoMovmientos 
         Caption         =   "Rastreo de Movimientos"
      End
      Begin VB.Menu mnuProComportamientoCta 
         Caption         =   "Comportamiento de Cuentas"
      End
   End
   Begin VB.Menu mnuMantenimiento 
      Caption         =   "&Mantenimiento"
      Begin VB.Menu mnuDivisas 
         Caption         =   "Divisas"
      End
      Begin VB.Menu mnuUnidadesNegocios 
         Caption         =   "Unidades de Negocios"
      End
      Begin VB.Menu mnuCentrosCostos 
         Caption         =   "Centros de Costos"
      End
      Begin VB.Menu mnuSepMant01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuTiposCuentas 
         Caption         =   "Tipos de Cuentas"
      End
      Begin VB.Menu mnuTiposAsientos 
         Caption         =   "Tipos de Asientos"
      End
      Begin VB.Menu mnuReportesSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPeriodos 
         Caption         =   "Periodos"
      End
      Begin VB.Menu mnuRepPeriodoFiscal 
         Caption         =   "Periodo Fiscal"
      End
      Begin VB.Menu mnuMantSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEmpresa 
         Caption         =   "Empresa"
      End
      Begin VB.Menu mnuRegistrosContabilidades 
         Caption         =   "Contabilidades"
      End
      Begin VB.Menu mnuMantSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfiguracionGeneral 
         Caption         =   "Configuración Estado Resultados"
      End
      Begin VB.Menu mnuConfeccionERPer 
         Caption         =   "Confección ER Personalizados"
      End
      Begin VB.Menu mnuMantSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInventarioPeriodicoAsiento 
         Caption         =   "Inventario Periodico (Asiento de Ajuste)"
      End
      Begin VB.Menu mnuInventarioPeriodico 
         Caption         =   "Inventario Periodico (Saldos)"
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
      Begin VB.Menu mnuAyudaSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "Acerca De.."
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
 
 StatusBar.Panels(2).Text = UCase(glogon.Usuario)
 StatusBar.Panels(6).Text = UCase(gCntX_Parametros.EmpresaLocal)

 txtAnio = gCntX_Parametros.PeriodoAnio
 txtMes = gCntX_Parametros.PeriodoMes
 tlbEmpresa.Buttons.Item(1).Caption = gCntX_Parametros.NombreEmpresa
  
 Me.Caption = App.ProductName & " - " & App.Major & "." & App.Minor & ".r" & App.Revision & GLOBALES.SysVersion

 frmContenedor.CD.HelpFile = App.HelpFile
 frmContenedor.CD.HelpCommand = cdlHelpContext 'cdlHelpContents
 If gCntX_Parametros.Explorador Then Call sbFormsCall("frmCntX_Explorer")
  
 Call Formularios(Me)
 Call RefrescaTags(Me)
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = 0 Then
      Cancel = True
      TimerSalir.Interval = 10
   End If
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
On Error GoTo vError
 sbCntX_Estado_Guarda
 Call sbLogonDSN(glogon.DSN, True)
 glogon.Conection.Close
 End
vError:
End Sub


Private Sub mnuCntAccionEditar_Click()
Dim frmX As Form
Call sbFormActivo("frmCntX_Explorer", frmX)
Call frmX.sbButtonPopUp(1)
End Sub

Private Sub mnuCntAccionBorrar_Click()
Dim frmX As Form
Call sbFormActivo("frmCntX_Explorer", frmX)
Call frmX.sbButtonPopUp(2)
End Sub


Private Sub mnuCntAccionesImprimir_Click()
Dim frmX As Form
Call sbFormActivo("frmCntX_Explorer", frmX)
Call frmX.sbButtonPopUp(4)
End Sub

Private Sub mnuCntAccionesMayorizar_Click()
Dim frmX As Form
Call sbFormActivo("frmCntX_Explorer", frmX)
Call frmX.sbButtonPopUp(5)
End Sub

Private Sub mnuCntAccionRefrescar_Click()
Dim frmX As Form
Call sbFormActivo("frmCntX_Explorer", frmX)
Call frmX.sbButtonPopUp(3)
End Sub

Private Sub mnuAcercaDe_Click()
 Call sbFormsCall("frmAcercaDe", 1, , , False)
End Sub


Private Sub mnuAreasTrabajoDefinicion_Click()
 Call sbClassCall("Contabilidad", 0, "frmCntX_AreaDefinicion")
End Sub

Private Sub mnuAreasTrabajoReportes_Click()
 Call sbClassCall("Contabilidad", 0, "frmCntX_AreasReportes")
End Sub


Private Sub mnuAutorizacionAsientos_Click()
 Call sbFormsCall("frmCntX_AsientosAutorizacion", 1, , , False)
End Sub

Private Sub mnuCambiaClave_Click()
Call sbFormsCall("frmCambiaClave", 1, , , False)
End Sub


Private Sub mnuCambioEmpresa_Click()
  gCntX_Parametros.MuestraTodas = True
  
  Call sbFormsCall("frmCntX_Seleccionar", 1, , , False)
  
  txtAnio.Text = gCntX_Parametros.PeriodoAnio
  txtMes.Text = gCntX_Parametros.PeriodoMes
  tlbEmpresa.Buttons.Item(1).Caption = gCntX_Parametros.NombreEmpresa
 
  frmCntX_Explorer.sbRefrescaArbol

End Sub

Private Sub mnuCambioSO_Click()
 Call sbFormsCall("frmLogonOS", 1, , , False)
 Call sbLogonDSN(glogon.DSN)

End Sub

Private Sub mnuCambioUsuario_Click()
Dim frmX As Form, pUsuario As String

'Guarda al Usuario actual, y si esta cambia (logea nuevamente y cierra ventanas abiertas)
pUsuario = Trim(UCase(glogon.Usuario))


'Cierra Formularios Abiertos
For Each frmX In Forms
   If Trim(frmX.Name) <> "MDIMenu" Then
        Unload frmX
   End If
Next

frmLogon.Show vbModal

Call sbgSEGInicializa

glogon.DSN = "ProGrX: Contabilidad"
Call sbLogonDSN(glogon.DSN)


Call sbCntX_ParametrosIniciales
  
gCntX_Parametros.MuestraTodas = False
frmCntX_Seleccionar.Show

Call MDIForm_Load

End Sub


Private Sub mnuCatalogoCuentas_Click()
 Call sbClassCall("Contabilidad", 0, "frmCntX_CatalogoCuentas")
End Sub

Private Sub mnuCentrosCostos_Click()
 Call sbClassCall("Contabilidad", 0, "frmCntX_CentrosCostos")
End Sub

Private Sub mnuCntX_ReporteBalanceGenRes_Click()
 Call sbFormsCall("frmCntX_RepGeneralResultados", 1, , , False)
End Sub

Private Sub mnuCntX_ReporteBalanceRsm_Click()
 Call sbFormsCall("frmCntX_RepRsmBalance", 1, , , False)

End Sub

Private Sub mnuConfeccionERPer_Click()
 Call sbFormsCall("frmCntX_EREspecial", 1, , , , Me)
End Sub

Private Sub mnuConfiguracionGeneral_Click()
  Call sbFormsCall("frmCntX_ERConfiguracion", 1, , , False)
End Sub

Private Sub mnuContenido_Click()
'   frmContenedor.CD.HelpCommand = cdlHelpContents
'   frmContenedor.CD.HelpCommand = cdlHelpContext
End Sub

Private Sub mnuCopiarEsquema_Click()
 Call sbClassCall("Contabilidad", 0, "frmCntX_Esquemas")
End Sub

Private Sub mnuDiferencialCambiario_Click()
 Call sbFormsCall("frmCntX_DivisasDC", 1, , , False)
End Sub

Private Sub mnuDiferidosCreacion_Click()
  Call sbClassCall("Contabilidad", 0, "frmCntX_DiferidosCreacion")
End Sub

Private Sub mnuDiferidosGeneracion_Click()
 Call sbFormsCall("frmCntX_DiferidosGeneracion", 1, , , False)
End Sub

Private Sub mnuDiferidosPlanillas_Click()
  Call sbClassCall("Contabilidad", 0, "frmCntX_DiferidosPlantilla")
End Sub


Private Sub mnuDivisas_Click()

  Call sbClassCall("Contabilidad", 0, "frmCntX_Divisas")

End Sub

Private Sub mnuEliminaAsientos_Click()
 Call sbFormsCall("frmCntX_UtilEliminaAsientos", 1, , , False)
End Sub

Private Sub mnuEliminaContabilidad_Click()
  Call sbClassCall("Contabilidad", 0, "frmCntX_UtilEliminaConta")
  
End Sub

Private Sub mnuEmpresa_Click()
   Call sbClassCall("Contabilidad", 0, "frmCntX_Empresa")
End Sub


Private Sub mnuExplorador_Click()

Call sbFormsCall("frmCntX_Explorer")

End Sub

Private Sub mnuInventarioPeriodico_Click()

  Call sbClassCall("Contabilidad", 0, "frmCntX_ERCuentas")
  
End Sub

Private Sub mnuInventarioPeriodicoAsiento_Click()

  Call sbClassCall("Contabilidad", 0, "frmCntX_AsientosInv")

End Sub

Private Sub mnuMayorización_Click()
 
Call sbFormsCall("frmCntX_MayorizacionFull", 1, , , False)

End Sub

Private Sub mnuPeriodos_Click()
  Call sbClassCall("Contabilidad", 0, "frmCntX_PeriodosDefinicion")
End Sub

Private Sub mnuPlantillaAsientoRate_Click()
  Call sbClassCall("Contabilidad", 0, "frmCntX_PlantillaRate")
End Sub


Private Sub mnuPlantillaAsientosRateGen_Click()
 Call sbFormsCall("frmCntX_PlantillaRateGen", 1, , , False)
End Sub

Private Sub mnuPlantillasDefinicion_Click()
  Call sbClassCall("Contabilidad", 0, "frmCntX_PlantillaAsientos")
End Sub

Private Sub mnuPlantillasGeneracion_Click()
 Call sbFormsCall("frmCntX_PlantillaAsientosGenera", 1, , , False)
End Sub

Private Sub mnuProcesosAdd_Click()
 Call sbFormsCall("frmCntX_ProcesosAdd", 1, , , False)
End Sub

Private Sub mnuProMezclas_Click()



End Sub

Private Sub mnuProRepConsolidados_Click()


End Sub

Private Sub mnuRazonesFinancieras_Click()
 Call sbClassCall("Contabilidad", 0, "frmCntX_RazonesFinanzas")
End Sub

Private Sub mnuRegistroAsientos_Click()
Call sbFormsCall("frmCntX_Asientos", , , , , Me)
End Sub

Private Sub mnuRegistrosContabilidades_Click()
Call sbFormsCall("frmCntX_Contabilidades", , , , False)
End Sub

Private Sub mnuRepBalanceComprobacionPosCierre_Click()
Call sbFormsCall("frmCntX_RepBalanceComprobacion", 1, , , , Me)
End Sub


Private Sub mnuRepBalanceSit_Click()
Call sbFormsCall("frmCntX_RepBalanceSituacion", 1, , , False)
End Sub

Private Sub mnuRepMovPeriodo_Click()
Call sbFormsCall("frmCntX_RepMovPeriodo", 1, , , , Me)
End Sub

Private Sub mnuRepMovTipoCuenta_Click()
 Call sbFormsCall("frmCntX_RepMovTipoCuenta", 1, , , , Me)
End Sub

Private Sub mnuRepMovTipoDocumento_Click()
 Call sbFormsCall("frmCntX_RepMovTipoDocumento", 1, , , , Me)
End Sub

Private Sub mnuReportePersonalizado_Click()
 Call sbFormsCall("frmCntX_RepEspeciales", 1, , , False)
End Sub

Private Sub mnuRepPeriodoFiscal_Click()
 Call sbClassCall("Contabilidad", 0, "frmCntX_Cierres")
End Sub

Private Sub mnuRestructurarMovimientosCuentas_Click()
Dim i As Integer, frmX As Form

i = MsgBox("Esta seguro que desea Restructurar los movimientos de este periodo...", vbYesNo)
If i = vbYes Then
   Set frmX = frmCntX_Procesos
   Call sbCntX_RestructuraMovimientosRSM(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes, frmX)
End If

End Sub

Private Sub mnuSalir_Click()
On Error Resume Next
 sbCntX_Estado_Guarda
 Call sbSEGCuentaLog("11")
 Call sbLogonDSN(glogon.DSN, True)
 glogon.Conection.Close
 End
End Sub

Private Sub mnuTipoDeCambio_Click()
 Call sbClassCall("Contabilidad", 0, "frmCntX_TipoCambioDefinicion")
End Sub

Private Sub mnuTiposAsientos_Click()
 Call sbClassCall("Contabilidad", 0, "frmCntX_TiposAsientos")
End Sub

Private Sub mnuTiposCuentas_Click()
 Call sbClassCall("Contabilidad", 0, "frmCntX_TiposCuentas")
End Sub


Private Sub sbCntX_Periodo_Refresh()
Dim strSQL As String, rs As New ADODB.Recordset
Dim strResultado As String

On Error GoTo vError

txtAnio = Val(txtAnio)
  
lblPeriodo.Text = fxCntX_PeriodoDesc(txtAnio, txtMes)
Call frmCntX_Explorer.sbRefrescaArbol


strSQL = "select estado from CntX_Periodos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and anio = " & txtAnio & " and mes = " & txtMes
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
 tlbPeriodo.Buttons.Item(1).ToolTipText = "Periodo No Definido"
 tlbPeriodo.Buttons.Item(1).Image = 5
 lblPeriodo.ForeColor = vbRed
Else
  If rs!estado = "P" Then
    tlbPeriodo.Buttons.Item(1).ToolTipText = "Periodo Pendiente"
    tlbPeriodo.Buttons.Item(1).Image = 4
    lblPeriodo.ForeColor = vbGrayText
  Else
    tlbPeriodo.Buttons.Item(1).ToolTipText = "Periodo Cerrado"
    tlbPeriodo.Buttons.Item(1).Image = 3
    lblPeriodo.ForeColor = vbBlack
  End If
End If
rs.Close
  
tlbPeriodo.Top = 0
Exit Sub

vError:

End Sub

Private Sub mnuUnidadesNegocios_Click()
 Call sbClassCall("Contabilidad", 0, "frmCntX_Unidades")
End Sub

Private Sub mnuVerificaAsientos_Click()
 Call sbFormsCall("frmCntX_UtilVerificaAsientos", 1, , , False)
End Sub


Private Sub timerSalir_Timer()

TimerSalir.Interval = 0
Call mnuSalir_Click

End Sub


Private Sub tlbCierre_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim iRespuesta As Integer, frmX As Form

Select Case ButtonMenu.Key
  Case "CierrePeriodo"
    iRespuesta = MsgBox("Esta seguro que desea Cerrar este periodo...", vbYesNo)
    If iRespuesta = vbYes Then
        'Reestructura Movimientos
        Set frmX = frmCntX_Procesos
        Call sbCntX_RestructuraMovimientosRSM(txtAnio.Text, txtMes.Text, frmX, False)
        
        'Cierra Periodo (Mensual)
        Me.MousePointer = vbHourglass
            Call sbCntX_PeriodoCierre(txtAnio.Text, txtMes.Text)
        Me.MousePointer = vbDefault
    End If
    
  Case "CierreFiscal"
    iRespuesta = MsgBox("Esta seguro que desea generar Asientos de Cierre Fiscal...", vbYesNo)
    If iRespuesta = vbYes Then
      Set frmX = frmCntX_Procesos
     'No se reestructuran los movimientos porque Para los Asientos de Cierre Fiscal, ya tuvo que realizar el cierre del periodo
     ' Call sbCntX_RestructuraMovimientosRSM(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes, frmX, False)
      Call sbCntX_CierreFiscal(frmX, txtMes.Text, txtAnio.Text)
    End If

End Select

End Sub

Private Sub tlbEmpresa_ButtonClick(ByVal Button As MSComctlLib.Button)
  gCntX_Parametros.MuestraTodas = True
  
  Call sbFormsCall("frmCntX_Seleccionar", 1, , , False)
  
  txtAnio = gCntX_Parametros.PeriodoAnio
  txtMes = gCntX_Parametros.PeriodoMes
  
  tlbEmpresa.Buttons.Item(1).Caption = gCntX_Parametros.NombreEmpresa

  Call frmCntX_Explorer.sbRefrescaArbol

End Sub

Private Sub tlbPeriodo_ButtonClick(ByVal Button As MSComctlLib.Button)
Call sbFormsCall("frmCntX_Periodos", 1, , , False)

txtMes.Text = gCntX_Parametros.PeriodoMes
txtAnio.Text = gCntX_Parametros.PeriodoAnio

End Sub

Private Sub tmr1_Timer()
 lblPeriodo.Refresh
 tmr1.Interval = 0
End Sub

Private Sub txtAnio_Change()
 Call sbCntX_Periodo_Refresh
End Sub

Private Sub txtEmpresa_Change()
 Call sbCntX_Periodo_Refresh
End Sub

Private Sub txtAnio_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
    gCntX_Parametros.PeriodoAnio = txtAnio.Text
vError:
End Sub

Private Sub txtMes_Change()
 Call sbCntX_Periodo_Refresh
End Sub

Private Sub txtMes_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtAnio.SetFocus
End Sub

Private Sub txtMes_KeyUp(KeyCode As Integer, Shift As Integer)
On Error GoTo vError
    gCntX_Parametros.PeriodoMes = txtMes.Text
vError:
End Sub
