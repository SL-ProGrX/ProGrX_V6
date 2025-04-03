VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.MDIForm MDIMenu 
   AutoShowChildren=   0   'False
   BackColor       =   &H00808080&
   Caption         =   "Activos Fijos"
   ClientHeight    =   5895
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9495
   Icon            =   "MDIMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerSalir 
      Left            =   1320
      Top             =   360
   End
   Begin Crystal.CrystalReport Crt 
      Left            =   600
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   120
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar BarraEstado 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   5580
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   2646
            MinWidth        =   2646
            TextSave        =   "02:43 p.m."
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            TextSave        =   "NUM"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnu_Sistema 
      Caption         =   "&Sistema"
      Begin VB.Menu mnuSeguridad 
         Caption         =   "Seguridad"
         Begin VB.Menu mnuCambioUsuario 
            Caption         =   "Cambio de Usuario"
         End
         Begin VB.Menu mnuCambioClave 
            Caption         =   "Cambio de Contraseña"
         End
         Begin VB.Menu mnuSeguridadSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCambioSO 
            Caption         =   "Cambio S.O."
         End
      End
      Begin VB.Menu sep_sistema_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Salir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
   Begin VB.Menu mnuActivos 
      Caption         =   "Activos"
      Begin VB.Menu mnu_Explorar 
         Caption         =   "Explorar"
      End
      Begin VB.Menu mnuActivoSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivosCatalogos 
         Caption         =   "Catálogos"
         Begin VB.Menu mnuTipoActivo 
            Caption         =   "Tipos de Activos"
         End
         Begin VB.Menu mnuActivosJustificacion 
            Caption         =   "Justificaciones"
         End
         Begin VB.Menu mnuActivosMotivosTraslados 
            Caption         =   "Motivos de Traslados"
         End
         Begin VB.Menu mnuEnlaceDepartamento 
            Caption         =   "Departamentos / Secciones"
         End
         Begin VB.Menu mnuEnlaceResponsables 
            Caption         =   "Personas"
         End
         Begin VB.Menu mnuEnlaceSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuEnlaceProveedores 
            Caption         =   "Lista de Proveedores"
         End
      End
      Begin VB.Menu mnuRegistroActivos 
         Caption         =   "Registro Activos"
      End
      Begin VB.Menu mnuReportesGenerales 
         Caption         =   "Reportes Generales"
      End
      Begin VB.Menu mnuActivosSep00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuObrasProceso 
         Caption         =   "Obras en Proceso"
         Visible         =   0   'False
         Begin VB.Menu mnuObrasTipos 
            Caption         =   "Tipos de Obras"
         End
         Begin VB.Menu mnuObrasTiposDesem 
            Caption         =   "Tipos de Desembolsos"
         End
         Begin VB.Menu mnuObrasProcSep0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuObrasProcesoRegistro 
            Caption         =   "Registro y Control"
         End
         Begin VB.Menu mnuObrasProcesoReportes 
            Caption         =   "Reportes"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuArrendamiento 
         Caption         =   "Arrendamiento Activos"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuActivosSep0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivosRenumeracion 
         Caption         =   "Renumeración de Activos"
      End
      Begin VB.Menu mnuActivosTraslados 
         Caption         =   "Cambio/Traslado de Responsable"
         Begin VB.Menu mnuActivosTrasladoPersonas 
            Caption         =   "Traslados entre Personas"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuActivosTrasladosEspacio1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuReAsignacion 
            Caption         =   "Traslado por Número de Placa"
         End
      End
      Begin VB.Menu mnuActivosPolizas 
         Caption         =   "Pólizas"
         Begin VB.Menu mnuPolizasRegistro 
            Caption         =   "Registro de Pólizas"
         End
         Begin VB.Menu mnuPolizasReportes 
            Caption         =   "Reportes de Pólizas"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuPolizasSep0 
            Caption         =   "-"
         End
         Begin VB.Menu mnuPolizasTipos 
            Caption         =   "Tipos de Pólizas"
         End
      End
      Begin VB.Menu mnuActivosSep01x 
         Caption         =   "-"
      End
      Begin VB.Menu AdicionRetiro 
         Caption         =   "Adiciones  / Retiros"
      End
      Begin VB.Menu mnuActivosRevaluacion 
         Caption         =   "Revaluaciones"
      End
      Begin VB.Menu mnuDeterioros 
         Caption         =   "Deterioros y desvalorizaciones"
      End
      Begin VB.Menu mnuActivoSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenerarAsientosDepreciacion 
         Caption         =   "Cierre del Periodo"
      End
      Begin VB.Menu mnuActivosTrasAsientos 
         Caption         =   "Traslado de Asientos"
      End
      Begin VB.Menu mnuActivoSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivoParametrosMenu 
         Caption         =   "Parámetros"
         Begin VB.Menu mnuActivosParametros 
            Caption         =   "Parámetros"
         End
         Begin VB.Menu mnuEnlacesParametros 
            Caption         =   "Parametros de Enlaces"
         End
      End
   End
   Begin VB.Menu mnuActivosExplorador 
      Caption         =   "Explorador: Activos Fijos"
      Visible         =   0   'False
      Begin VB.Menu mnuActivosAccionNuevo 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu mnuActivosAccionPropiedades 
         Caption         =   "Propiedades"
      End
      Begin VB.Menu mnuActivosAccionEliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu sep_activos_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivosAccionDepreciacion 
         Caption         =   "Depreciación"
      End
      Begin VB.Menu Sep_Activos_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivosAccionActualizar 
         Caption         =   "Actualizar"
      End
      Begin VB.Menu mnuActivosAccionImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu sep_E_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivosAccionCerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu mnu_A_Activos 
      Caption         =   "Activos"
      Visible         =   0   'False
      Begin VB.Menu mnu_A_Nuevo 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu mnu_A_Guardar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu mnu_A_Eliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu mnu_A_Actualizar 
         Caption         =   "Actualizar"
      End
      Begin VB.Menu mnu_A_Imprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu sep_A_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_A_Cerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu mnu_Adicion_Retiro 
      Caption         =   "AdicionRetiro"
      Visible         =   0   'False
      Begin VB.Menu mnu_AR_Nuevo 
         Caption         =   "Nuevo"
      End
      Begin VB.Menu mnu_AR_Guardar 
         Caption         =   "Guardar"
      End
      Begin VB.Menu mnu_AR_Eliminar 
         Caption         =   "Eliminar"
      End
      Begin VB.Menu sep_AR_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_AR_Imprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnu_AR_Actualizar 
         Caption         =   "Actualizar"
      End
      Begin VB.Menu sep_AR_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_AR_Cerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&?"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contenido"
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "&Buscar por ayuda en..."
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAcercaDe 
         Caption         =   "&Acerca de..."
      End
   End
End
Attribute VB_Name = "MDIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Const NAME_COLUMN = 0
Const TYPE_COLUMN = 1
Const SIZE_COLUMN = 2
Const DATE_COLUMN = 3

Private Sub AdicionRetiro_Click()
   Call sbClassCall("Activos", 0, "frmActivos_AdicionRetiro")
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = 0 Then
      Cancel = True
      TimerSalir.Interval = 10
   End If
End Sub




Private Sub mnuActivosAccionNuevo_Click()
Dim frmX As Form
Call sbFormActivo("frmActivos_Explorador", frmX)
Call frmX.Tlb_Herramientas_ButtonClick(frmX.tlb_Herramientas.Buttons(1))
End Sub

Private Sub mnuActivosAccionPropiedades_Click()
Dim frmX As Form
Call sbFormActivo("frmActivos_Explorador", frmX)
Call frmX.Tlb_Herramientas_ButtonClick(frmX.tlb_Herramientas.Buttons(2))
End Sub

Private Sub mnuActivosAccionEliminar_Click()
Dim frmX As Form
Call sbFormActivo("frmActivos_Explorador", frmX)
Call frmX.Tlb_Herramientas_ButtonClick(frmX.tlb_Herramientas.Buttons(3))
End Sub


Private Sub mnuActivosAccionDepreciacion_Click()
Dim frmX As Form
Call sbFormActivo("frmActivos_Explorador", frmX)
Call frmX.Tlb_Herramientas_ButtonClick(frmX.tlb_Herramientas.Buttons(5))

End Sub

Private Sub mnuActivosAccionActualizar_Click()
Dim frmX As Form
Call sbFormActivo("frmActivos_Explorador", frmX)
Call frmX.Tlb_Herramientas_ButtonClick(frmX.tlb_Herramientas.Buttons(7))
End Sub


Private Sub mnuActivosAccionImprimir_Click()
Dim frmX As Form
Call sbFormActivo("frmActivos_Explorador", frmX)
Call frmX.Tlb_Herramientas_ButtonClick(frmX.tlb_Herramientas.Buttons(8))
End Sub

Private Sub mnu_Explorar_Click()
    frmActivos_Explorador.Show
End Sub

Private Sub mnuAcercaDe_Click()
frmAcercaDe.Show vbModal
End Sub

Private Sub mnuActivosJustificacion_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Justificaciones")
End Sub

Private Sub mnuActivosMotivosTraslados_Click()
    Call sbClassCall("Activos", 0, "frmActivos_TrasladosMotivos")
End Sub

Private Sub mnuActivosParametros_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Parametros")
End Sub

Private Sub mnuActivosRenumeracion_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Renumeracion")
End Sub

Private Sub mnuActivosRevaluacion_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Revaluaciones")
End Sub

Private Sub mnuActivosTrasAsientos_Click()
    Call sbClassCall("Activos", 0, "frmActivos_TrasladoAsientos")
End Sub

Private Sub mnuActivosTrasladoPersonas_Click()
    Call sbClassCall("Activos", 0, "frmActivos_ResponsablesCambio")
End Sub

Private Sub mnuArrendamiento_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Arrendamientos")
End Sub

Private Sub mnuCambioClave_Click()
 frmCambiaClave.Show vbModal
End Sub

Private Sub mnuCambioSO_Click()
frmLogonOS.Show vbModal
Call sbLogonDSN(glogon.DSN)
End Sub

Private Sub mnuCambioUsuario_Click()

 Call Main
 
End Sub

Private Sub mnuDeterioros_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Deterioros")
End Sub

Private Sub mnuEnlaceDepartamento_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Departamentos")
End Sub

Private Sub mnuEnlaceProveedores_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Proveedores")
End Sub

Private Sub mnuEnlaceResponsables_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Personas")
End Sub


Private Sub mnuEnlacesParametros_Click()
    Call sbClassCall("Activos", 0, "frmActivos_ParametrosEnlaces")
End Sub

Private Sub mnuGenerarAsientosDepreciacion_Click()
    Call sbClassCall("Activos", 0, "frmActivos_CierrePeriodo")
End Sub

Private Sub mnuObrasProcesoRegistro_Click()
    Call sbClassCall("Activos", 0, "frmActivos_ObrasProceso")
End Sub

Private Sub mnuObrasProcesoReportes_Click()
    Call sbClassCall("Activos", 0, "frmActivos_ObrasReportes")
End Sub

Private Sub mnuObrasTipos_Click()
    Call sbClassCall("Activos", 0, "frmActivos_ObrasTipos")
End Sub

Private Sub mnuObrasTiposDesem_Click()
    Call sbClassCall("Activos", 0, "frmActivos_ObrasTipoDesem")
End Sub

Private Sub mnuPolizasRegistro_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Polizas")
End Sub

Private Sub mnuPolizasReportes_Click()
    Call sbClassCall("Activos", 0, "frmActivos_PolizasReportes")
End Sub

Private Sub mnuPolizasTipos_Click()
    Call sbClassCall("Activos", 0, "frmActivos_PolizasTipos")
End Sub

Private Sub mnuReAsignacion_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Reasignacion")
End Sub

Private Sub mnuRegistroActivos_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Main")
End Sub

Private Sub mnuReportesGenerales_Click()
    Call sbClassCall("Activos", 0, "frmActivos_Reportes")
End Sub

Private Sub mnuTipoActivo_Click()
    Call sbClassCall("Activos", 0, "frmActivos_TiposActivo")
End Sub

Private Sub mnuHelpContents_Click()
  On Error Resume Next
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
  If Err Then
    MsgBox Err.Description
  End If
End Sub

Private Sub mnuHelpSearch_Click()
  On Error Resume Next
  
  Dim nRet As Integer
  nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
  If Err Then
    MsgBox Err.Description
  End If
End Sub


Private Sub MDIForm_Load()
'Carga Fondo de la Empresa
Set Me.Picture = fxImagen_Leer("select Fondo_Pantalla from SIF_Empresa", "Fondo_Pantalla")

If Me.Picture = 0 Then
   Set Me.Picture = LoadPicture(App.Path & "\SifFondoDefault.jpg")
End If

 
 frmActivos_Explorador.Show
 BarraEstado.Panels(1).Text = "US: " & glogon.Usuario
 BarraEstado.Panels(2).Text = "SR: " & UCase(glogon.Servidor)
 BarraEstado.Panels(3).Text = "DB: " & UCase(glogon.BaseDatos)
 MDIMenu.Caption = App.ProductName & " - " & App.Major & "." & App.Minor & "." & App.Revision & GLOBALES.SysVersion
End Sub

Private Sub mnu_A_Cerrar_Click()
'    Unload frmActivos_Main
End Sub

Private Sub mnu_E_Cerrar_Click()
    Unload frmActivos_Explorador
End Sub

Private Sub mnu_Salir_Click()
Call sbLogonDSN(glogon.DSN, True)
End
End Sub

Private Sub TimerSalir_Timer()
TimerSalir.Interval = 0
Call mnu_Salir_Click
End Sub
