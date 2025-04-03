VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIMenu 
   BackColor       =   &H8000000C&
   Caption         =   "Menu"
   ClientHeight    =   4725
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   7815
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   4470
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "10:27 AM"
            Object.ToolTipText     =   "Hora"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Usuario Activo"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Servidor de Coneccion"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   3704
            MinWidth        =   3704
            Object.ToolTipText     =   "Base de Datos"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "29/04/2006"
            Object.ToolTipText     =   "Fecha del Sistema"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   989
            MinWidth        =   989
            TextSave        =   "CAPS"
            Object.ToolTipText     =   "Tecla de CAP activa"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   988
            MinWidth        =   988
            TextSave        =   "NUM"
            Object.ToolTipText     =   "NumLock"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3177
            MinWidth        =   3177
            Object.ToolTipText     =   "Fecha Proceso Créditos"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3177
            MinWidth        =   3177
            Object.ToolTipText     =   "Fecha Proceso Ahorros"
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuSeguridad 
         Caption         =   "Seguridad"
         Begin VB.Menu mnuCambioUsuario 
            Caption         =   "Cambio de Usuario"
         End
         Begin VB.Menu mnuCambioContrasena 
            Caption         =   "Cambio Contraseña"
         End
         Begin VB.Menu mnuSeguridadSep1 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCambioSO 
            Caption         =   "Cambio de S.O."
         End
      End
      Begin VB.Menu mnuArchivoSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuModulos 
      Caption         =   "&Módulos"
      Begin VB.Menu mnuCargaPlanillas 
         Caption         =   "Carga de Planillas"
      End
      Begin VB.Menu mnuModulosSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlantillasCreditos 
         Caption         =   "Plantillas de Créditos"
      End
      Begin VB.Menu mnuModulosSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConciliacionSaldos 
         Caption         =   "Conciliación de Saldos"
      End
   End
   Begin VB.Menu mnuAcercaDe 
      Caption         =   "&Acerca de"
   End
End
Attribute VB_Name = "MDIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
 StatusBar.Panels(2).Text = "US: " & glogon.Usuario
 StatusBar.Panels(3).Text = "SR: " & UCase(glogon.Servidor)
 StatusBar.Panels(4).Text = "DB: " & UCase(glogon.BaseDatos)
 
 MDIMenu.Caption = App.ProductName & " - " & App.Major & "." & App.Minor & ".r" & App.Revision

End Sub

Private Sub mnuAcercaDe_Click()
 frmAcercaDe.Show vbModal
End Sub

Private Sub mnuCambioContrasena_Click()
 frmCambiaClave.Show vbModal
End Sub

Private Sub mnuCambioSO_Click()
 frmLogonOS.Show vbModal
 Call sbXLogonOSInit(glogon.DSN)
End Sub

Private Sub mnuCambioUsuario_Click()
 frmLogon.Show vbModal
 Call sbCargaParametros
 Call sbXLogonOSInit(glogon.DSN)
End Sub

Private Sub mnuCargaPlanillas_Click()
 Call MuestraForms(frmSIFCargaPlanilla)
End Sub

Private Sub mnuConciliacionSaldos_Click()
 frmSIFConciliacionSaldos.Show
End Sub

Private Sub mnuPlantillasCreditos_Click()
 Call MuestraForms(frmSIFCargaCreditos)
End Sub

Private Sub mnuSalir_Click()
On Error Resume Next
 Call sbXLogonOSEnd(glogon.DSN)
 glogon.Conection.Close
 End
End Sub
