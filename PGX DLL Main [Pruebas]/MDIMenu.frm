VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.MDIForm MDIMenu 
   BackColor       =   &H8000000C&
   Caption         =   "Menu: Pruebas de DLLs y Modulos SIF"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList imgMain 
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
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":6862
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":D0C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":13926
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":1A188
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIMenu.frx":1A295
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   2775
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "11:58 p.m."
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
            TextSave        =   "02/12/2009"
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
            Enabled         =   0   'False
            Object.Width           =   988
            MinWidth        =   988
            TextSave        =   "NUM"
            Object.ToolTipText     =   "NumLock"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3177
            MinWidth        =   3177
            Object.ToolTipText     =   "Fecha Proceso"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6705
            MinWidth        =   6705
            Object.ToolTipText     =   "Oficina"
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
   Begin MSComctlLib.Toolbar tlbConfiguracion 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Style           =   1
      ImageList       =   "imgMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "clave"
            Object.ToolTipText     =   "Cambio de Contraseña"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "usuario"
            Object.ToolTipText     =   "Cambia de usuario "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SO"
            Object.ToolTipText     =   "Cambia de Sistema Operativo"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir de la Aplicación"
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin VB.Image Image1 
         Height          =   615
         Left            =   9240
         Picture         =   "MDIMenu.frx":1A3CC
         Stretch         =   -1  'True
         Top             =   -720
         Width           =   375
      End
   End
End
Attribute VB_Name = "MDIMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Activate()
 StatusBar.Panels(2).Text = "US: " & glogon.Usuario
 StatusBar.Panels(3).Text = "SR: " & UCase(glogon.Servidor)
 StatusBar.Panels(4).Text = "DB: " & UCase(glogon.BaseDatos)
 
 StatusBar.Panels(8).Text = "FP: " & Format(GLOBALES.glngFechaCR, "####-##")
 StatusBar.Panels(9).Text = "OF: " & GLOBALES.gOficina

   
 Me.Caption = App.ProductName & " - " & App.Major & "." & App.Minor & ".r" & App.Revision
 frmContenedor.CD.HelpFile = App.HelpFile
 frmContenedor.CD.HelpCommand = cdlHelpContext 'cdlHelpContents

End Sub

Private Sub MDIForm_Load()
 frmMain.Show
End Sub

Private Sub tlbConfiguracion_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "clave"
         frmCambiaClave.Show vbModal
  Case "usuario"
        frmLogon.Show vbModal
        Call sbSIFParametrosInicializa
        Call sbLogonDSN(glogon.DSN)
  Case "SO"
        frmLogonOS.Show vbModal
        Call sbLogonDSN(glogon.DSN)
  Case "salir"
        On Error Resume Next
            Call sbLogonDSN(glogon.DSN, True)
            glogon.Conection.Close
         End
End Select


        
        
End Sub
