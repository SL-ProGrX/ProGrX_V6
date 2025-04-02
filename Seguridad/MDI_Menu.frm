VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "codejock.controls.v22.1.0.ocx"
Object = "{B8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.TaskPanel.v22.1.0.ocx"
Begin VB.MDIForm MDIMenu 
   BackColor       =   &H00C0C0C0&
   Caption         =   "..."
   ClientHeight    =   9120
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   18315
   HelpContextID   =   1001
   Icon            =   "MDI_Menu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin XtremeTaskPanel.TaskPanel tpContabilidad 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   18315
      _Version        =   1441793
      _ExtentX        =   32306
      _ExtentY        =   741
      _StockProps     =   64
      VisualTheme     =   13
      ItemLayout      =   2
      HotTrackStyle   =   1
      Begin XtremeSuiteControls.PushButton btnCliente 
         Height          =   360
         Index           =   0
         Left            =   0
         TabIndex        =   3
         ToolTipText     =   "Seleccione un Cliente"
         Top             =   0
         Width           =   1455
         _Version        =   1441793
         _ExtentX        =   2566
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Cliente"
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FlatStyle       =   -1  'True
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "MDI_Menu.frx":08CA
      End
      Begin XtremeSuiteControls.FlatEdit txtCliente 
         Height          =   315
         Left            =   1560
         TabIndex        =   2
         Top             =   60
         Width           =   10065
         _Version        =   1441793
         _ExtentX        =   17754
         _ExtentY        =   556
         _StockProps     =   77
         ForeColor       =   0
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Locked          =   -1  'True
         Appearance      =   6
         UseVisualStyle  =   0   'False
      End
   End
   Begin VB.Timer Timer_Load 
      Interval        =   5
      Left            =   120
      Top             =   840
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8865
      Width           =   18315
      _ExtentX        =   32306
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Bevel           =   0
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "01:59:p. m."
            Object.ToolTipText     =   "Hora"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   989
            MinWidth        =   989
            TextSave        =   "MAYÚS"
            Object.ToolTipText     =   "Tecla de CAP activa"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Bevel           =   0
            Object.Width           =   988
            MinWidth        =   988
            TextSave        =   "NÚM"
            Object.ToolTipText     =   "NumLock"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
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
      Begin VB.Menu mnuArchivoSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSeguridadSbMnu 
         Caption         =   "Seguridad"
         Begin VB.Menu mnuCambiaUsuario 
            Caption         =   "Cambia de Usuario"
         End
         Begin VB.Menu mnuCambiaClave 
            Caption         =   "Cambia Clave"
         End
         Begin VB.Menu munSeguridadSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuSegCambioSO 
            Caption         =   "Cambio S.O."
         End
      End
      Begin VB.Menu mnuArchivoSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuAcciones 
      Caption         =   "&Acciones"
      Visible         =   0   'False
      Begin VB.Menu mnuAccionEditar 
         Caption         =   "Editar"
      End
      Begin VB.Menu mnuAccionReporte 
         Caption         =   "Reportes"
      End
      Begin VB.Menu mnuAccionPermisos 
         Caption         =   "Permisos"
      End
      Begin VB.Menu mnuAccionSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAccionCerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu mnuSeguridad 
      Caption         =   "&Seguridad"
      Begin VB.Menu mnuUsuarios 
         Caption         =   "&Usuarios"
      End
      Begin VB.Menu mnuGrupos 
         Caption         =   "&Roles"
      End
      Begin VB.Menu mnuAdministradores 
         Caption         =   "Administradores"
      End
      Begin VB.Menu mnuSegXPerSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPermisos 
         Caption         =   "Permisos por Roles"
      End
      Begin VB.Menu mnuPermisosXOpcion 
         Caption         =   "Permisos x Opción"
      End
      Begin VB.Menu mnuSeguridadSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSegOpciones 
         Caption         =   "&Opciones"
         Begin VB.Menu mnuModulos 
            Caption         =   "&Módulos"
         End
         Begin VB.Menu mnuFormularios 
            Caption         =   "&Formularios"
         End
         Begin VB.Menu mnuSegSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuOpciones 
            Caption         =   "&Opciones"
         End
         Begin VB.Menu mnuSegSep1_2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuMenu 
            Caption         =   "Menú de Acceso"
         End
         Begin VB.Menu mnuBE_Movimientos 
            Caption         =   "Bitácora Especial: Movimientos"
         End
      End
      Begin VB.Menu mnuSeguridadSeparador3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuControlAplicacines 
         Caption         =   "Control de Aplicaciones"
      End
      Begin VB.Menu mnuSeguridadSeparador4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBitacoraSeguridad 
         Caption         =   "Bitácora de Seguridad"
      End
      Begin VB.Menu mnuBitacora 
         Caption         =   "Bitácora de Movimientos"
      End
      Begin VB.Menu mnuSeguridadSeparador5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParametros 
         Caption         =   "Parámetros"
      End
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnuClientesSeguridad 
         Caption         =   "Clientes: Usuarios y Roles"
      End
      Begin VB.Menu mnuConfiguracionSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuServicios 
         Caption         =   "Servicios"
      End
      Begin VB.Menu mnuVendedores 
         Caption         =   "Vendedores"
      End
      Begin VB.Menu mnuConfiguracionSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConfigDistribucionPolitica 
         Caption         =   "Distribución Politica"
      End
      Begin VB.Menu mnuConfigTiposIDs 
         Caption         =   "Tipos de Identificación"
      End
      Begin VB.Menu mnuConfigClienteClasifica 
         Caption         =   "Clasificación de Clientes"
      End
      Begin VB.Menu mnuSepradorApps1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAppEstadistica 
         Caption         =   "Apps: Estadistica"
         Begin VB.Menu mnuAppEstadisticaSub 
            Caption         =   "Estadistica Apps"
            Index           =   0
         End
         Begin VB.Menu mnuAppEstadisticaSub 
            Caption         =   "Tipos de Movimientos"
            Index           =   1
         End
         Begin VB.Menu mnuAppEstadisticaSub 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuAppEstadisticaSub 
            Caption         =   "Sincronizar Webs/Apps"
            Index           =   3
         End
      End
      Begin VB.Menu mnuSepradorApps2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLoadFile 
         Caption         =   "Carga de Archivos"
      End
      Begin VB.Menu mnuDbExterna 
         Caption         =   "Consulta Base de Datos Externa (Migración)"
      End
   End
   Begin VB.Menu mnuLA 
      Caption         =   "Administración"
      Begin VB.Menu mnuLA_Usuarios 
         Caption         =   "Usuarios"
      End
      Begin VB.Menu mnuLA_Permisos 
         Caption         =   "Permisos"
      End
      Begin VB.Menu mnuLA_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLA_BitacoraCuentas 
         Caption         =   "Bitácora de Cuentas"
      End
      Begin VB.Menu mnuLA_BitacoraMovimientos 
         Caption         =   "Bitácora de Movimientos"
      End
      Begin VB.Menu mnuLA_Informes 
         Caption         =   "Informes"
      End
      Begin VB.Menu mnuLA_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLA_AppsEstadisticas 
         Caption         =   "Estadisticas Apps"
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
      Begin VB.Menu mnuSeparadorAyuda1 
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

Private Sub btnCliente_Click(Index As Integer)

Call sbFormsCall("frmPGX_ClienteSelect", vbModal, , , False, Me)
txtCliente.Tag = gPortal.Empresa_Id
txtCliente.Text = gPortal.Empresa_Name

Call sbAdmin_Rols_Load

End Sub

Private Sub MDIForm_Load()
vModulo = 13

Me.Caption = App.ProductName & " [ " & App.Major & "." & App.Minor & "." & App.Revision & ".r" & GLOBALES.SysVersion & " ]"

Me.BackColor = RGB(78, 111, 178)

mnuSeguridad.Visible = False
mnuConfiguracion.Visible = False
mnuLA.Visible = False

If Sys_Portal_Admin_Valid(glogon.Usuario) Then
    mnuSeguridad.Visible = True
    mnuConfiguracion.Visible = True
Else
    mnuLA.Visible = True
End If

frmUS_Explorer.Show


If glogon.AppStatus = 1 Then
   MsgBox "Actualización disponible!", vbInformation
'   Call sbFormsCall("frmCC_AppStatus", , , , False, Me)
End If

End Sub

Private Sub MDIForm_Resize()
On Error Resume Next
txtCliente.Width = Me.Width - (txtCliente.Left + 200)
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
End
End Sub

Private Sub mnuAccionEditar_Click()
Call frmUS_Explorer.sbButtonPopUp(1)
End Sub

Private Sub mnuAccionPermisos_Click()
Call frmUS_Explorer.sbButtonPopUp(3)
End Sub

Private Sub mnuAccionReporte_Click()
Call frmUS_Explorer.sbButtonPopUp(2)
End Sub

Private Sub mnuAcercaDe_Click()
 frmAcercaDe.Show vbModal
End Sub

Private Sub mnuAdministradores_Click()
Call sbFormsCall("frmUS_Admin_Rol", vbModal)
End Sub

Private Sub mnuAppEstadisticaSub_Click(Index As Integer)
Select Case Index
  Case 0 'Estadisticas
     Call sbFormsCall("frmSYS_APP_LOG", 0, , , True, Me)
  Case 1 'Tipos de Hits
     Call sbFormsCall("frmSYS_APP_HITS", vbModal, , , True, Me)
  Case 3 'Sincronizar WebApps
     Call sbFormsCall("frmSYS_Portal_WebApp_Sincroniza", vbModal, , , True, Me)
  
End Select
End Sub

Private Sub mnuBE_Movimientos_Click()
 Call sbFormsCall("frmUS_BE_TiposMov", 0, , , True, Me)
End Sub

Private Sub mnuBitacora_Click()
 Call sbFormsCall("frmUS_Bitacora", 0, , , True, Me)
End Sub

Private Sub mnuBitacoraSeguridad_Click()
Call sbFormsCall("frmUS_CuentaLog", vbModal, , , True, Me)
End Sub

Private Sub mnuCambiaClave_Click()
 frmCambiaClave.Show vbModal
End Sub

Private Sub mnuCambiaUsuario_Click()
 frmLogon.Show vbModal
 
' Call sbXLogonOSInit(glogon.DSN)

End Sub

Private Sub mnuClientes_Click()
  frmPGX_Clientes.Show
End Sub

Private Sub mnuClientesSeguridad_Click()
 
 Unload frmUS_Roles_Membrecias
 Call sbFormsCall("frmUS_Roles_Membrecias", vbModal, , , True, Me)
End Sub

Private Sub mnuConfigClienteClasifica_Click()
  Call sbFormsCall("frmPGX_ClientesClasifica", 0, , , True, Me)
End Sub

Private Sub mnuConfigDistribucionPolitica_Click()
  Call sbFormsCall("frmPGX_DistribucionPolitica", 0, , , True, Me)
End Sub

Private Sub mnuConfigTiposIDs_Click()
  Call sbFormsCall("frmPGX_ClientesTiposIDs", 0, , , True, Me)
End Sub

Private Sub mnuContenido_Click()
   frmContenedor.CD.HelpCommand = cdlHelpContents
   frmContenedor.CD.ShowHelp
   frmContenedor.CD.HelpCommand = cdlHelpContext
End Sub

Private Sub mnuControlAccesos_Click()

 Call sbFormsCall("frmUS_Accesos", vbModal, , , True, Me)
End Sub

Private Sub mnuControlAplicacines_Click()
 Call sbFormsCall("frmUS_Aplicaciones", vbModal, , , True, Me)
End Sub

Private Sub mnuDbExterna_Click()
 Call sbFormsCall("frmSYS_BD_Analisis", 1, , , True)
  
End Sub

Private Sub mnuExplorador_Click()
frmUS_Explorer.Show
End Sub

Private Sub mnuFormularios_Click()
 Call sbFormsCall("frmUS_Formularios", 1, , , True)
End Sub

Private Sub mnuGrupos_Click()
 Call sbFormsCall("frmUS_Roles", 1, , , True)
End Sub

Private Sub mnuLA_AppsEstadisticas_Click()
     Call sbFormsCall("frmSYS_APP_LOG", 0, , , True, Me)
End Sub

Private Sub mnuLA_BitacoraCuentas_Click()
Call sbFormsCall("frmUS_CuentaLog", vbModal, , , True, Me)
End Sub

Private Sub mnuLA_BitacoraMovimientos_Click()
 Call sbFormsCall("frmUS_Bitacora", 0, , , True, Me)
End Sub

Private Sub mnuLA_Informes_Click()
 Call sbFormsCall("frmUS_ReporteUsuarios", vbModal, , , True, Me)
End Sub

Private Sub mnuLA_Permisos_Click()
 Unload frmUS_Roles_Membrecias
 Call sbFormsCall("frmUS_Roles_Membrecias", vbModal, , , True, Me)
End Sub

Private Sub mnuLA_Usuarios_Click()
  Call sbFormsCall("frmUS_Usuarios", 0, , , True, Me)
End Sub

Private Sub mnuLoadFile_Click()
Call sbFormsCall("frmSYS_Load_File", 0, , , True, Me)
End Sub

Private Sub mnuMenu_Click()
 Call sbFormsCall("frmUS_Menus", 1, , , True)
End Sub

Private Sub mnuModulos_Click()
 Call sbFormsCall("frmUS_Modulos", 1, , , True)
End Sub

Private Sub mnuOpciones_Click()
 frmUS_Opciones.Show
End Sub

Private Sub mnuParametros_Click()
frmUS_Parametros.Show vbModal
End Sub

Private Sub mnuPermisos_Click()
frmUS_DerechosNew.Show
End Sub

Private Sub mnuPermisosXOpcion_Click()
frmUS_DerechosXOpcion.Show
End Sub

Private Sub mnuSalir_Click()
 End
End Sub

Private Sub mnuSegCambioSO_Click()
 frmLogonOS.Show vbModal
' Call sbXLogonOSInit(glogon.DSN)
End Sub

Private Sub mnuServicios_Click()
  Call sbFormsCall("frmPGX_Servicios", 0, , , True, Me)
End Sub

Private Sub mnuUsuarios_Click()
  Call sbFormsCall("frmUS_Usuarios", 0, , , True, Me)
End Sub

Private Sub mnuVendedores_Click()
  Call sbFormsCall("frmPGX_Vendedores", 0, , , True, Me)
End Sub

Private Sub Timer_Load_Timer()


Dim strSQL As String, rs As New ADODB.Recordset
Dim pLoad As Boolean

On Error GoTo vErrorTimerLoad

pLoad = False

If Timer_Load.Interval <> 60000 Then
   pLoad = True
End If

Timer_Load.Interval = 60000

Call sbAdmin_Rols_Load

If gPortal.Empresa_Id = -1 And Not gAdminAccess.Admin_Portal Then
    Call btnCliente_Click(0)
End If

Exit Sub

vErrorTimerLoad:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub
