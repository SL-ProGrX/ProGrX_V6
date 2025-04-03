VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmCR_SeguimientoTramitesOld 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seguimiento de Trámites"
   ClientHeight    =   6870
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8235
   HelpContextID   =   3027
   Icon            =   "frmCR_SeguimientoTramitesOld.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6870
   ScaleWidth      =   8235
   Begin VB.Frame fraUser 
      Height          =   2655
      Left            =   1680
      TabIndex        =   96
      Top             =   840
      Visible         =   0   'False
      Width           =   4575
      Begin VB.Image imgConsultaUsers 
         Height          =   255
         Left            =   4080
         Picture         =   "frmCR_SeguimientoTramitesOld.frx":030A
         Stretch         =   -1  'True
         ToolTipText     =   "Cierra Consulta"
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Usuarios Tramitadores de la Operación Activa"
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   105
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblTesoreria 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1560
         TabIndex        =   104
         Top             =   1965
         Width           =   2535
      End
      Begin VB.Label lblFormaliza 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1560
         TabIndex        =   103
         Top             =   1605
         Width           =   2535
      End
      Begin VB.Label lblResoluciona 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1560
         TabIndex        =   102
         Top             =   1245
         Width           =   2535
      End
      Begin VB.Label lblRecibe 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1560
         TabIndex        =   101
         Top             =   885
         Width           =   2535
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tesoreria"
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   3
         Left            =   360
         TabIndex        =   100
         Top             =   1965
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Formaliza"
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   2
         Left            =   360
         TabIndex        =   99
         Top             =   1605
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Resoluciona"
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   98
         Top             =   1245
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Recibe"
         ForeColor       =   &H8000000E&
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   97
         Top             =   885
         Width           =   1215
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4920
      Top             =   240
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
            Picture         =   "frmCR_SeguimientoTramitesOld.frx":0614
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramitesOld.frx":0938
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCR_SeguimientoTramitesOld.frx":1218
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   91
      Top             =   0
      Width           =   8235
      _ExtentX        =   14526
      _ExtentY        =   688
      BandCount       =   2
      _CBWidth        =   8235
      _CBHeight       =   390
      _Version        =   "6.7.8988"
      Child1          =   "tlbPrincipal"
      MinHeight1      =   330
      Width1          =   4260
      NewRow1         =   0   'False
      Child2          =   "tlbAux"
      MinHeight2      =   330
      Width2          =   2520
      NewRow2         =   0   'False
      Begin MSComctlLib.Toolbar tlbAux 
         Height          =   330
         Left            =   4455
         TabIndex        =   93
         Top             =   30
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   582
         ButtonWidth     =   2064
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   4
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Cálculo"
               Key             =   "calculo"
               Object.ToolTipText     =   "Cálculo de la Operación"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "PreAnalisis"
               Key             =   "preanalisis"
               Object.ToolTipText     =   "PreAnalisis de la Solicitud"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Usuarios"
               Key             =   "usersx"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbPrincipal 
         Height          =   330
         Left            =   165
         TabIndex        =   92
         Top             =   30
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Style           =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "nuevo"
               Object.ToolTipText     =   "Inserta (Agrega) un registro nuevo a la Base de Datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "editar"
               Object.ToolTipText     =   "Modifica (Edita) el registro en pantalla"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "guardar"
               Object.ToolTipText     =   "Guarda la información del registro en la base de datos"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "deshacer"
               Object.ToolTipText     =   "Deshace toda modificación realizada recientemente en el registro actual"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "reportes"
               Object.ToolTipText     =   "Imprime el listado seleccionado"
               Object.Tag             =   "1"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "RepActas"
                     Text            =   "Actas"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "RepPreAnalisis"
                     Text            =   "Pre Analisis"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "RepGarantia"
                     Text            =   "Garantía"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "repBoleta"
                     Text            =   "Boleta"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ayuda"
               Object.ToolTipText     =   "Ayuda General"
               Object.Tag             =   "1"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "cerrar"
               Object.ToolTipText     =   "Cierra esta ventana"
               Object.Tag             =   "1"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fraConsulta 
      Height          =   3660
      Left            =   0
      TabIndex        =   83
      Top             =   960
      Visible         =   0   'False
      Width           =   8340
      Begin VB.TextBox txtConNombre 
         Height          =   315
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   85
         ToolTipText     =   "Presiones F4 para Consultar"
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox txtConCedula 
         Height          =   315
         Left            =   960
         TabIndex        =   84
         ToolTipText     =   "Presiones F4 para Consultar"
         Top             =   240
         Width           =   1695
      End
      Begin MSComctlLib.ListView lswBusca 
         Height          =   2940
         Left            =   120
         TabIndex        =   87
         Top             =   600
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5186
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   16711680
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "operacion"
            Text            =   "#Operación"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "codigo"
            Text            =   "Código"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Key             =   "cedula"
            Text            =   "Cédula"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Key             =   "fecha"
            Text            =   "Fecha Solicitud"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Key             =   "monto"
            Text            =   "Monto Sol"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Estado Sol"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Estado EC"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Proceso"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Image imgConsultaCerrar 
         Height          =   255
         Left            =   7680
         Picture         =   "frmCR_SeguimientoTramitesOld.frx":1AF2
         Stretch         =   -1  'True
         ToolTipText     =   "Cierra Consulta"
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Cédula"
         Height          =   255
         Index           =   26
         Left            =   120
         TabIndex        =   86
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.TextBox txtOperacion 
      Height          =   315
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   1695
   End
   Begin TabDlg.SSTab ssTabOperacion 
      Height          =   2175
      Left            =   0
      TabIndex        =   1
      Top             =   4680
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   3836
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Recepción"
      TabPicture(0)   =   "frmCR_SeguimientoTramitesOld.frx":1DFC
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraRecepcion"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Resolución"
      TabPicture(1)   =   "frmCR_SeguimientoTramitesOld.frx":1E18
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label17(0)"
      Tab(1).Control(1)=   "Label16"
      Tab(1).Control(2)=   "Label1(5)"
      Tab(1).Control(3)=   "Label1(17)"
      Tab(1).Control(4)=   "Label1(18)"
      Tab(1).Control(5)=   "Label1(20)"
      Tab(1).Control(6)=   "imgRecibir"
      Tab(1).Control(7)=   "Label17(1)"
      Tab(1).Control(8)=   "imgRequisitos"
      Tab(1).Control(9)=   "Label1(27)"
      Tab(1).Control(10)=   "Line5"
      Tab(1).Control(11)=   "Line6"
      Tab(1).Control(12)=   "medMontoAprobado"
      Tab(1).Control(13)=   "dtpFechaFormalizacion"
      Tab(1).Control(14)=   "dtpFechaRes"
      Tab(1).Control(15)=   "chkPrimera"
      Tab(1).Control(16)=   "txtCuotaAprobada"
      Tab(1).Control(17)=   "txtInteresAprobado"
      Tab(1).Control(18)=   "txtPlazoAprobado"
      Tab(1).Control(19)=   "fraEstadoResolucion"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "Formalización"
      TabPicture(2)   =   "frmCR_SeguimientoTramitesOld.frx":1E34
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label10(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "imgGuardaFecDesembolso"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Line1"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Line2"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Label17(2)"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "dtpDesembolso"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "imgFormalizacion"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "tlbFormalizacion"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "fraEstadoFormalizacion"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "txtPagare"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "txtAno"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "cboMes"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "chkEnviarATesoreria"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "txtDocumento"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "cboRecursos"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).ControlCount=   17
      TabCaption(3)   =   "Estado de Cuenta"
      TabPicture(3)   =   "frmCR_SeguimientoTramitesOld.frx":1E50
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(21)"
      Tab(3).Control(1)=   "Label1(22)"
      Tab(3).Control(2)=   "Label1(23)"
      Tab(3).Control(3)=   "Label1(24)"
      Tab(3).Control(4)=   "Label1(25)"
      Tab(3).Control(5)=   "chkDeducirPlanilla"
      Tab(3).Control(6)=   "txtSaldo"
      Tab(3).Control(7)=   "txtInteresPagado"
      Tab(3).Control(8)=   "txtAmortizado"
      Tab(3).Control(9)=   "txtEstadoActual"
      Tab(3).Control(10)=   "txtProceso"
      Tab(3).Control(11)=   "cmdDeduccionPlanilla"
      Tab(3).ControlCount=   12
      Begin VB.ComboBox cboRecursos 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   107
         Top             =   1680
         Width           =   3735
      End
      Begin VB.TextBox txtDocumento 
         Height          =   315
         Left            =   3840
         MaxLength       =   18
         TabIndex        =   79
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton cmdDeduccionPlanilla 
         Caption         =   "&Aplicar"
         Height          =   975
         Left            =   -68280
         Picture         =   "frmCR_SeguimientoTramitesOld.frx":1E6C
         Style           =   1  'Graphical
         TabIndex        =   82
         Top             =   480
         Width           =   975
      End
      Begin VB.CheckBox chkEnviarATesoreria 
         Alignment       =   1  'Right Justify
         Caption         =   "Enviar a Tesoreria"
         Height          =   195
         Left            =   960
         TabIndex        =   78
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Frame fraEstadoResolucion 
         Height          =   1575
         Left            =   -69600
         TabIndex        =   74
         Top             =   360
         Width           =   2535
         Begin VB.OptionButton optResolucion 
            Caption         =   "&Aprobar"
            ForeColor       =   &H00FF0000&
            Height          =   495
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton optResolucion 
            Caption         =   "&Denegar"
            ForeColor       =   &H000000FF&
            Height          =   495
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdAplicaResolucion 
            Caption         =   "&Aplicar"
            Height          =   975
            Left            =   1440
            Picture         =   "frmCR_SeguimientoTramitesOld.frx":2176
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame fraRecepcion 
         Height          =   1335
         Left            =   -74880
         TabIndex        =   37
         Top             =   480
         Width           =   7815
         Begin VB.CommandButton cmdCargos 
            Caption         =   "&Cargos Adicionales"
            Height          =   975
            Left            =   6360
            Picture         =   "frmCR_SeguimientoTramitesOld.frx":2480
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton cmdRequisitos 
            Caption         =   "&Requisitos"
            Height          =   975
            Left            =   5160
            Picture         =   "frmCR_SeguimientoTramitesOld.frx":278A
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdFirmas 
            Caption         =   "&Firmas"
            Height          =   975
            Left            =   3960
            Picture         =   "frmCR_SeguimientoTramitesOld.frx":2A94
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdCalculoOperacion 
            Caption         =   "&Cálculo Operación"
            Height          =   975
            Left            =   2760
            Picture         =   "frmCR_SeguimientoTramitesOld.frx":2D9E
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdDatosPersonales 
            Caption         =   "&Datos Personales"
            Height          =   975
            Left            =   1560
            Picture         =   "frmCR_SeguimientoTramitesOld.frx":30A8
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdFiadores 
            Caption         =   "&Fiadores"
            Height          =   975
            Left            =   360
            Picture         =   "frmCR_SeguimientoTramitesOld.frx":33B2
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.TextBox txtProceso 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70440
         TabIndex        =   34
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtEstadoActual 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -70440
         TabIndex        =   33
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox txtAmortizado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73440
         TabIndex        =   31
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtInteresPagado 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73440
         TabIndex        =   30
         Top             =   840
         Width           =   1575
      End
      Begin VB.TextBox txtSaldo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   -73440
         TabIndex        =   29
         Top             =   480
         Width           =   1575
      End
      Begin VB.CheckBox chkDeducirPlanilla 
         Alignment       =   1  'Right Justify
         Caption         =   "Deducir Cuotas por Planilla"
         Height          =   375
         Left            =   -70440
         TabIndex        =   28
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox cboMes 
         Height          =   315
         ItemData        =   "frmCR_SeguimientoTramitesOld.frx":36BC
         Left            =   3840
         List            =   "frmCR_SeguimientoTramitesOld.frx":36E7
         TabIndex        =   23
         ToolTipText     =   "Mes a procesar"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtAno 
         Height          =   315
         Left            =   5160
         TabIndex        =   22
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtPagare 
         Height          =   315
         Left            =   960
         MaxLength       =   18
         TabIndex        =   20
         Top             =   840
         Width           =   1695
      End
      Begin VB.Frame fraEstadoFormalizacion 
         Height          =   1335
         Left            =   5720
         TabIndex        =   16
         Top             =   720
         Width           =   2295
         Begin VB.CommandButton cmdAplicarFormalizacion 
            Caption         =   "&Aplicar"
            Height          =   975
            Left            =   1200
            Picture         =   "frmCR_SeguimientoTramitesOld.frx":3750
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optFormalizacion 
            Caption         =   "&Anular"
            ForeColor       =   &H000000FF&
            Height          =   495
            Index           =   1
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton optFormalizacion 
            Caption         =   "&Formalizar"
            ForeColor       =   &H00FF0000&
            Height          =   495
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin MSComctlLib.Toolbar tlbFormalizacion 
         Height          =   360
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   635
         ButtonWidth     =   2355
         ButtonHeight    =   582
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgFormalizacion"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Créditos"
               Key             =   "refundiciones"
               Object.ToolTipText     =   "Refundiciones de Créditos"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Desembolsos"
               Key             =   "desembolsos"
               Object.ToolTipText     =   "Desembolsos Internos y Externos"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Retenciones"
               Key             =   "retenciones"
               Object.ToolTipText     =   "Refundiciones de Retenciones Fijas"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Firmas"
               Key             =   "firmas"
               Object.ToolTipText     =   "Registro de Firmas"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Requisitos"
               Key             =   "requisitos"
               Object.ToolTipText     =   "Activación de Requisitos"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cargos "
               Key             =   "cargos"
               Object.ToolTipText     =   "Activación de Cargos "
               ImageIndex      =   4
            EndProperty
         EndProperty
         BorderStyle     =   1
      End
      Begin VB.TextBox txtPlazoAprobado 
         Height          =   315
         Left            =   -74160
         TabIndex        =   10
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox txtInteresAprobado 
         Height          =   315
         Left            =   -72960
         TabIndex        =   9
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txtCuotaAprobada 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   -74160
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CheckBox chkPrimera 
         Alignment       =   1  'Right Justify
         Caption         =   "Deducir Primer Cuota"
         Height          =   195
         Left            =   -71640
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker dtpFechaRes 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   -71040
         TabIndex        =   4
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59768835
         CurrentDate     =   36306
      End
      Begin MSComCtl2.DTPicker dtpFechaFormalizacion 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "MM/dd/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   -71040
         TabIndex        =   5
         ToolTipText     =   "Fecha para formalizar el préstamo"
         Top             =   840
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59768835
         CurrentDate     =   36115
      End
      Begin MSMask.MaskEdBox medMontoAprobado 
         Height          =   315
         Left            =   -74160
         TabIndex        =   36
         Top             =   480
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "###,###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSComctlLib.ImageList imgFormalizacion 
         Left            =   7560
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCR_SeguimientoTramitesOld.frx":3A5A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCR_SeguimientoTramitesOld.frx":3D7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCR_SeguimientoTramitesOld.frx":465E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCR_SeguimientoTramitesOld.frx":4982
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmCR_SeguimientoTramitesOld.frx":526A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComCtl2.DTPicker dtpDesembolso 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         Height          =   315
         Left            =   3840
         TabIndex        =   94
         Top             =   1680
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59768835
         CurrentDate     =   36306
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         X1              =   -74400
         X2              =   -69600
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   -74400
         X2              =   -69600
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label17 
         Caption         =   "Recursos y Fecha de Desembolso"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   95
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   5640
         X2              =   240
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   5640
         X2              =   240
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Label Label1 
         Caption         =   "Activar Requisitos"
         Height          =   255
         Index           =   27
         Left            =   -74160
         TabIndex        =   106
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Image imgRequisitos 
         Height          =   255
         Left            =   -72600
         Picture         =   "frmCR_SeguimientoTramitesOld.frx":558E
         Stretch         =   -1  'True
         ToolTipText     =   "Actualiza Requisitos"
         Top             =   1800
         Width           =   255
      End
      Begin VB.Image imgGuardaFecDesembolso 
         Height          =   255
         Left            =   5400
         Picture         =   "frmCR_SeguimientoTramitesOld.frx":5898
         Stretch         =   -1  'True
         ToolTipText     =   "Guarda la Fecha de Desembolso"
         Top             =   1680
         Width           =   255
      End
      Begin VB.Label Label17 
         Caption         =   "Poner como Recibida"
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   1
         Left            =   -71640
         TabIndex        =   90
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Image imgRecibir 
         Height          =   255
         Left            =   -69960
         Picture         =   "frmCR_SeguimientoTramitesOld.frx":5BA2
         Stretch         =   -1  'True
         ToolTipText     =   "Poner Como Recibida Nuevamente"
         Top             =   1800
         Width           =   255
      End
      Begin VB.Label Label4 
         Caption         =   "# Doc. Ref."
         Height          =   255
         Left            =   2760
         TabIndex        =   80
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Proceso"
         Height          =   255
         Index           =   25
         Left            =   -71760
         TabIndex        =   35
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Estado Actual"
         Height          =   255
         Index           =   24
         Left            =   -71760
         TabIndex        =   32
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Amortizado"
         Height          =   255
         Index           =   23
         Left            =   -74760
         TabIndex        =   27
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Interes Pagado"
         Height          =   255
         Index           =   22
         Left            =   -74760
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Saldo"
         Height          =   255
         Index           =   21
         Left            =   -74760
         TabIndex        =   25
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Pri.Deduc"
         Height          =   255
         Left            =   2760
         TabIndex        =   24
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label10 
         Caption         =   "# Pagaré"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cuota"
         Height          =   255
         Index           =   20
         Left            =   -74880
         TabIndex        =   14
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Interes"
         Height          =   255
         Index           =   18
         Left            =   -73560
         TabIndex        =   13
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Plazo"
         Height          =   255
         Index           =   17
         Left            =   -74880
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Monto"
         Height          =   255
         Index           =   5
         Left            =   -74880
         TabIndex        =   11
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label16 
         Caption         =   "Formalización"
         Height          =   255
         Left            =   -72120
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label17 
         Caption         =   "Resolución"
         Height          =   255
         Index           =   0
         Left            =   -72120
         TabIndex        =   6
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame fraOperacion 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3615
      Left            =   0
      TabIndex        =   41
      Top             =   960
      Width           =   8295
      Begin VB.ComboBox cboDestino 
         Height          =   315
         ItemData        =   "frmCR_SeguimientoTramitesOld.frx":5EAC
         Left            =   3720
         List            =   "frmCR_SeguimientoTramitesOld.frx":5EAE
         Style           =   2  'Dropdown List
         TabIndex        =   109
         Top             =   840
         Width           =   4455
      End
      Begin VB.TextBox txtCedula 
         Height          =   315
         Left            =   960
         TabIndex        =   57
         ToolTipText     =   "Presiones F4 para Consultar"
         Top             =   120
         Width           =   1695
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   56
         ToolTipText     =   "Presiones F4 para Consultar"
         Top             =   120
         Width           =   4455
      End
      Begin VB.TextBox txtCodigo 
         Height          =   315
         Left            =   960
         TabIndex        =   55
         ToolTipText     =   "Presiones F4 para Consultar"
         Top             =   480
         Width           =   1695
      End
      Begin VB.TextBox txtDescripcion 
         Enabled         =   0   'False
         Height          =   315
         Left            =   3720
         TabIndex        =   54
         ToolTipText     =   "Presiones F4 para Consultar"
         Top             =   480
         Width           =   4455
      End
      Begin VB.TextBox txtPlazoSolicitado 
         Height          =   315
         Left            =   3720
         TabIndex        =   53
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtInteresSolicitado 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   0
         EndProperty
         Height          =   315
         Left            =   5160
         TabIndex        =   52
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtCuotaSolicitada 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   5130
            SubFormatType   =   1
         EndProperty
         Enabled         =   0   'False
         Height          =   315
         Left            =   6360
         MultiLine       =   -1  'True
         TabIndex        =   51
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox cboGarantia 
         Height          =   315
         ItemData        =   "frmCR_SeguimientoTramitesOld.frx":5EB0
         Left            =   960
         List            =   "frmCR_SeguimientoTramitesOld.frx":5EB2
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   1680
         Width           =   1695
      End
      Begin VB.ComboBox cboEstado 
         Height          =   315
         ItemData        =   "frmCR_SeguimientoTramitesOld.frx":5EB4
         Left            =   6360
         List            =   "frmCR_SeguimientoTramitesOld.frx":5EB6
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   1680
         Width           =   1815
      End
      Begin VB.ComboBox cboComite 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   2040
         Width           =   4215
      End
      Begin VB.ComboBox cboTipoDocumento 
         Height          =   315
         ItemData        =   "frmCR_SeguimientoTramitesOld.frx":5EB8
         Left            =   6360
         List            =   "frmCR_SeguimientoTramitesOld.frx":5EBA
         TabIndex        =   46
         Text            =   "cboTipoDocumento"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.ComboBox cboBanco 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   2400
         Width           =   4215
      End
      Begin VB.TextBox txtCuentaAhorros 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6360
         TabIndex        =   44
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox txtObservaciones 
         Height          =   885
         Left            =   960
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   2760
         Width           =   7215
      End
      Begin MSMask.MaskEdBox medMontoSolicitado 
         Height          =   315
         Left            =   960
         TabIndex        =   42
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Format          =   "###,###,###,##0.00"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dtpFechaSolicitud 
         Height          =   315
         Left            =   3720
         TabIndex        =   49
         Top             =   1680
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59768835
         CurrentDate     =   36434
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   8160
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         X1              =   8160
         X2              =   120
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Label Label1 
         Caption         =   "Destino"
         Height          =   255
         Index           =   28
         Left            =   2760
         TabIndex        =   108
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Cédula"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   73
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   2
         Left            =   2760
         TabIndex        =   72
         Top             =   120
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Código"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   71
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Descripción"
         Height          =   255
         Index           =   4
         Left            =   2760
         TabIndex        =   70
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Garantía"
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   69
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Estado"
         Height          =   255
         Index           =   10
         Left            =   5400
         TabIndex        =   68
         Top             =   1680
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   11
         Left            =   2760
         TabIndex        =   67
         Top             =   1680
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Evaluado "
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   66
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Emitir"
         Height          =   255
         Index           =   13
         Left            =   5400
         TabIndex        =   65
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta Aho."
         Height          =   255
         Index           =   14
         Left            =   5400
         TabIndex        =   64
         Top             =   2400
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Banco"
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   63
         Top             =   2400
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Cuota"
         Height          =   255
         Index           =   8
         Left            =   5880
         TabIndex        =   62
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Interes"
         Height          =   255
         Index           =   7
         Left            =   4440
         TabIndex        =   61
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Plazo"
         Height          =   255
         Index           =   6
         Left            =   2760
         TabIndex        =   60
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Monto"
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   59
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Observac."
         Height          =   255
         Index           =   19
         Left            =   120
         TabIndex        =   58
         Top             =   2760
         Width           =   1095
      End
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   8160
      X2              =   120
      Y1              =   880
      Y2              =   880
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   120
      X2              =   8160
      Y1              =   880
      Y2              =   880
   End
   Begin VB.Image imgConsulta 
      Height          =   255
      Left            =   2760
      Picture         =   "frmCR_SeguimientoTramitesOld.frx":5EBC
      Stretch         =   -1  'True
      ToolTipText     =   "Consulta de Operaciones"
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "Operación"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frmCR_SeguimientoTramitesOld"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vMensaje As String 'Envia Mensajes en Fallas de Verificacion
Dim vEdita As Boolean 'Indica si se esta actualizando o insertando
Dim vPasaFormalizacion As Boolean 'Indica si una formalizacion normal se procesa o no
Dim vDocumentoFormalizacion As Boolean 'Indica si se debe de generar una nota de debito
'Por incluir una formalizacion que no pasa a Tesoreria o el Monto Girado es Cero

Private Sub cboBanco_Click()
txtCuentaAhorros = fxCuentaAhorros(txtCedula, cboBanco.ItemData(cboBanco.ListIndex))
End Sub

Private Sub cboBanco_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
  txtCuentaAhorros = fxCuentaAhorros(txtCedula, cboBanco.ItemData(cboBanco.ListIndex))
  txtObservaciones.SetFocus
End If
End Sub

Private Sub cboComite_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboTipoDocumento.SetFocus
End Sub


Private Sub cboDestino_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then medMontoSolicitado.SetFocus
End Sub

Private Sub cboEstado_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboComite.SetFocus
End Sub

Private Sub cboGarantia_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 If KeyCode = vbKeyReturn Then dtpFechaSolicitud.SetFocus
End Sub

Private Sub cboTipoDocumento_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cboBanco.SetFocus
End Sub


Private Sub sbEnvioTesoreria(vGenerado As Boolean, vOperacion As Long)
Dim strSQL As String, rs As New ADODB.Recordset, vFecha As Date
Dim lngSol As Long, rsTmp As New ADODB.Recordset
'Utilizar este procedimiento y sus referencias solo para Creditos Rapidos

vFecha = Format(fxFechaServidor, "yyyy/mm/dd")

strSQL = "select R.id_solicitud,R.codigo,R.cedula,S.nombre,R.monto_girado,R.documento_referido" _
       & ",C.ctaPuente,R.monto_girado,R.cod_banco,R.cta_banco,R.emitir,B.ctaconta" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " inner join Catalogo C on R.codigo = C.codigo" _
       & " inner join Bancos B on R.cod_banco = B.id_banco" _
       & " where R.id_solicitud = " & vOperacion
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs!monto_girado = 0 Or rs!monto_girado < 0 Then
  rs.Close
  Exit Sub
End If

If vGenerado Then
    strSQL = "insert cheques(id_banco,tipo,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
           & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza" _
           & ",fecha_emision,ndocumento) values(" & rs!Cod_Banco & ",'" & rs!emitir & "','" _
           & rs!Cedula & "','" & rs!Nombre & "'," & rs!monto_girado & ",'" & Format(vFecha, "yyyy/mm/dd") _
           & "','I','P','CC','C','" & rs!cta_banco & "','" & rs!id_solicitud & "','" & rs!Codigo & "',0," _
           & rs!id_solicitud & ",'S','S','" & Format(vFecha, "yyyy/mm/dd") & "','" _
           & Mid(Trim(rs!documento_referido), 4, 30) & "')"

Else
    strSQL = "insert cheques(id_banco,tipo,codigo,beneficiario,monto,fecha_solicitud,estado,estadoi" _
           & ",modulo,submodulo,cta_ahorros,detalle1,detalle2,referencia,op,genera,actualiza) values(" & rs!Cod_Banco _
           & ",'" & rs!emitir & "','" & rs!Cedula & "','" & rs!Nombre & "'," & rs!monto_girado _
           & ",'" & Format(vFecha, "yyyy/mm/dd") & "','P','P','CC','C','" & rs!cta_banco _
           & "','" & rs!id_solicitud & "','" & rs!Codigo & "',0," & rs!id_solicitud & ",'S','S')"
End If
glogon.Conection.Execute strSQL

'Recupera Consecutivo Tesoreria
strSQL = "select max(nsolicitud) as Solicitud from cheques where codigo ='" & rs!Cedula _
      & "' and op = " & rs!id_solicitud
rsTmp.Open strSQL, glogon.Conection, adOpenStatic
  lngSol = rsTmp!solicitud
rsTmp.Close


'Cancela la cuenta del pasivo o puente cargada en el asiento de formalizacion
strSQL = "insert ck_detalle(nsolicitud,cuenta_contable,monto,debehaber,linea) values(" _
       & lngSol & ",'" & Trim(rs!ctapuente) & "'," & rs!monto_girado & ",'D',1)"
glogon.Conection.Execute strSQL
'Bancos
strSQL = "insert ck_detalle(nsolicitud,cuenta_contable,monto,debehaber,linea) values(" _
       & lngSol & ",'" & Trim(rs!ctaconta) & "'," & rs!monto_girado & ",'H',2)"
glogon.Conection.Execute strSQL
rs.Close


End Sub

Private Sub sbTramiteRapido(vPrimerCuota As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim curDias As Currency, curRefunde As Currency, curMonto As Currency
Dim curIntCuota As Currency, curAmortizaCuota As Currency, FechaUltima As Long
Dim vFecha As Date, vRefundeSaldo As Currency, iMes As Integer, lngAnio As Long
Dim vPaso As Boolean, Porcentaje As Double, vPriDeduc As Long
Dim vFechaCalculo As Date, curPSD As Currency

On Error GoTo vError

vFecha = fxFechaServidor
curPSD = 0

strSQL = "update reg_creditos set fechaforp = '" & Format(vFecha, "yyyy/mm/dd") _
       & "',cod_grupo = '" & fxCodigoCbo(cboRecursos) _
       & "' where id_solicitud = " & Operacion.Operacion
glogon.Conection.Execute strSQL

dtpFechaFormalizacion.Value = vFecha

curDias = fxInteresesHastaFormalizar

strSQL = "select cuota_poliza" _
       & " from reg_creditos where id_solicitud =" & Operacion.Operacion
rs.Open strSQL, glogon.Conection, adOpenStatic
  curPSD = IIf(IsNull(rs!cuota_poliza), 0, rs!cuota_poliza)
rs.Close


If vPrimerCuota = 1 Then
  
  vPriDeduc = fxPrimerDeduccionCuota
  curIntCuota = CCur(medMontoAprobado.Text) * CCur(txtInteresAprobado) / 1200
  curAmortizaCuota = CCur(txtCuotaAprobada) - curIntCuota
  FechaUltima = GLOBALES.glngFechaCR

  'Fecha de Ultima Deducción
  iMes = Month(vFecha)
  lngAnio = Year(vFecha)
  
     If iMes = 12 Then
        iMes = 1
        lngAnio = lngAnio + 1
     Else
        iMes = iMes + 1
     End If
     'Calcular Intereses Hasta el Ultimo día del Mes
     vFechaCalculo = CDate(lngAnio & "/" & Format(iMes, "00") & "/01")
     vFechaCalculo = DateAdd("d", -1, vFechaCalculo)
     
     curDias = ((CCur(txtInteresAprobado) / 36000) * CCur(medMontoAprobado.Text) * (Abs(DateDiff("d", vFechaCalculo, vFecha)) + 1))
     
  FechaUltima = CLng((lngAnio & Format(iMes, "00")))


Else
  
  vFechaCalculo = fxFechaCalculo()
    
  curDias = ((CCur(txtInteresAprobado) / 36000) * CCur(medMontoAprobado.Text) * (Abs(DateDiff("d", vFechaCalculo, vFecha)) + 1))
  
  curIntCuota = 0
  curAmortizaCuota = 0
  FechaUltima = fxFechaProcesoAnterior
  vPriDeduc = fxPrimerDeduccion
End If



'Inicia Transacciones
glogon.Conection.BeginTrans

'Nuevo Esquema para Que Abone a la Morosidad de Todas las Operaciones

vRefundeSaldo = 0

strSQL = "delete refundiciones where id_solicitudR = " & Operacion.Operacion
glogon.Conection.Execute strSQL

strSQL = "delete refunde_retencion where id_solicitudR = " & Operacion.Operacion
glogon.Conection.Execute strSQL

strSQL = "select R.id_solicitud,R.codigo,R.saldo,C.Retencion,C.poliza" _
       & ",coalesce(sum(M.intc),0) as IntCor,coalesce(sum(M.intM),0) as IntMor" _
       & ",coalesce(sum(M.amortiza),0) as Amortiza" _
       & " from reg_creditos R  inner join Catalogo C on R.codigo = C.codigo" _
       & " left join Morosidad M on R.id_solicitud = M.id_solicitud" _
       & " and M.estado = 'A' Where R.estado = 'A' and R.proceso = 'N' and R.saldo > 0" _
       & " and R.cedula = '" & Operacion.Cedula & "' group by  R.id_solicitud,R.codigo,R.saldo,C.Retencion,C.poliza"
rs.Open strSQL, glogon.Conection, adOpenForwardOnly
Do While Not rs.EOF
 If Trim(UCase(rs!retencion)) = "S" Or Trim(UCase(rs!poliza)) = "S" Then
  'RETENCIONES
  If (rs!Amortiza + rs!IntCor + rs!IntMor) > 0 Then
    strSQL = "insert refunde_retencion(id_solicitudr,codigor,id_solicitud,codigo,monto,mora,fecha,saldo_anterior) " _
           & "values(" & Operacion.Operacion & ",'" & Operacion.Codigo & "'," & rs!id_solicitud & ",'" _
           & rs!Codigo & "',0," & (rs!IntCor + rs!IntMor + rs!Amortiza) & ",'" _
           & Format(vFecha, "yyyy/mm/dd") & "',0)"
    glogon.Conection.Execute strSQL
    vRefundeSaldo = vRefundeSaldo + (rs!Amortiza + rs!IntCor + rs!IntMor)
  End If
 
 Else
  'CARTERA - SOLO SE VERIFICA EL CODIGO AQUI, PORQUE EL S.T. SOLO ACEPTA CARTERA
  'PERO SI PUEDE REFUNDIR RETENCIONES
  
    If UCase(Trim(rs!Codigo)) = UCase(Trim(Operacion.Codigo)) Then
      'Cancelar Saldo y Morosidad
       strSQL = "insert refundiciones(ID_SOLICITUDR,CODIGOR,ID_SOLICITUD,CODIGO,MONTO,INTCOR,INTMOR,FECHA)" _
              & " values(" & Operacion.Operacion & ",'" & Operacion.Codigo & "'," & rs!id_solicitud _
              & ",'" & Trim(rs!Codigo) & "'," & rs!Saldo & "," & rs!IntCor & "," & rs!IntMor _
              & ",'" & Format(vFecha, "yyyy/mm/dd") & "')"
       glogon.Conection.Execute strSQL
       vRefundeSaldo = vRefundeSaldo + (rs!Saldo + rs!IntCor + rs!IntMor)
    
    Else
      
      'Cancelar Solo Morosidad
      If (rs!Amortiza + rs!IntCor + rs!IntMor) > 0 Then
        strSQL = "insert refundiciones(ID_SOLICITUDR,CODIGOR,ID_SOLICITUD,CODIGO,MONTO,INTCOR,INTMOR,FECHA)" _
               & " values(" & Operacion.Operacion & ",'" & Operacion.Codigo & "'," & rs!id_solicitud _
               & ",'" & Trim(rs!Codigo) & "'," & rs!Amortiza & "," & rs!IntCor & "," & rs!IntMor _
               & ",'" & Format(vFecha, "yyyy/mm/dd") & "')"
        glogon.Conection.Execute strSQL
        vRefundeSaldo = vRefundeSaldo + (rs!Amortiza + rs!IntCor + rs!IntMor)
      End If
    End If
 End If
 
 rs.MoveNext
Loop
rs.Close


vPaso = True

'Verifica que no se pase del 90% si es un prestamo sobre ahorros
'Se incluye para Verificar el 90% del ahorro
If UCase(cboGarantia.Text) = "SOBRE AHORROS" Then
    strSQL = "select CR_POR_AHORRO  from par_ahcr"
    rs.Open strSQL, glogon.Conection, adOpenStatic
      Porcentaje = rs!Cr_Por_ahorro
    rs.Close
    
    strSQL = "select coalesce(sum(saldo),0) as Saldos from reg_creditos where estado = 'A'" _
           & " and saldo > 0 and garantia = 'A' and cedula = '" & Operacion.Cedula & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    curRefunde = rs!saldos 'Saldos Sobre Ahorros
    rs.Close
         
    
    strSQL = "select coalesce(AHORRO,0) as Ahorro, coalesce(Capitaliza,0) as Capital" _
           & " from ahorro_consolidado where cedula = '" & Operacion.Cedula & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    If ((CCur(medMontoAprobado.Text) - vRefundeSaldo) + curRefunde) > ((rs!Ahorro + rs!capital) * (Porcentaje / 100)) Then
       rs.Close
       'HAY QUE DARLE VUELTA A LA RESOLUCION
       glogon.Conection.RollbackTrans
       
       strSQL = "update reg_creditos set estadosol = 'R' where id_solicitud = " & Operacion.Operacion _
              & " and estadosol in('R','P','A','D')"
       glogon.Conection.Execute strSQL
       
       MsgBox "El monto aprobado más los saldos que quedan pendientes en sobre ahorros son mayores" & vbCrLf _
             & " al porcentaje reglamentario...", vbCritical
       Exit Sub
    End If
    rs.Close
End If
'Fin de Verificacion sobre ahorros

'On Error GoTo vError
'
''Inicia Transacciones
'glogon.Conection.BeginTrans

'BITACORA DE LA RESOLUCION
Call Bitacora("Registra", "Resolucion Aprobada a la OP: " & Operacion.Operacion)


'Formaliza Operacion en Tramite

strSQL = "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
       & "FECHAP,TCON,NCON,ESTADO,ESTADO_ASIENTO) values('" & Operacion.Codigo & "'," _
       & Operacion.Operacion & "," & CCur(txtCuotaAprobada) & "," & CCur(txtCuotaAprobada) + curDias _
       & "," & curIntCuota + curDias & "," & curAmortizaCuota & ",'" & Format(vFecha, "yyyy/mm/dd") _
       & "'," & GLOBALES.glngFechaCR & ",2," & Operacion.Operacion & ",'A','G')"
glogon.Conection.Execute strSQL

strSQL = "update reg_creditos set tesoreria = '" & Format(vFecha, "yyyy/mm/dd") & "'," _
       & "userfor='" & glogon.Usuario & "',usertesoreria = '" & glogon.Usuario & "'," _
       & "firma_deudor=1,estadosol='F',estado='A',interesc =" & curDias + curIntCuota & "," _
       & "amortiza =" & curAmortizaCuota & ", saldo = montoapr - " & curAmortizaCuota _
       & ", prideduc = " & vPriDeduc & ",fecult=" & FechaUltima _
       & ", Fecha_Calculo_int = '" & Format(vFechaCalculo, "yyyy/mm/dd") & "'" _
       & ",monto_girado= montoapr - " & (vRefundeSaldo + curAmortizaCuota + curDias + curIntCuota + curPSD) _
       & ",pagare =" & Operacion.Operacion _
       & ",cuotas_planilla = " & vPrimerCuota _
       & ",documento_referido = '" & fxTipoDocumento(cboTipoDocumento) & "-" & Operacion.Documento & "'" _
       & ",saldo_mes = montoapr - " & curAmortizaCuota _
       & " where id_solicitud = " & Operacion.Operacion
glogon.Conection.Execute strSQL

'Abonar Refundiciones
strSQL = "select * from refundiciones where id_solicitudr = " & Operacion.Operacion
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  Call sbAbonaRefundiciones(rs!id_solicitud, rs!Codigo, rs!Monto, rs!IntCor, rs!IntMor, vFecha)
  rs.MoveNext
Loop
rs.Close

'Abonar Retenciones
strSQL = "select * from refunde_retencion where id_solicitudr = " & Operacion.Operacion
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  Call sbAbonaRetenciones(rs!id_solicitud, rs!Codigo, rs!Monto, rs!mora, vFecha)
  rs.MoveNext
Loop
rs.Close

           
'Crea Asiento de Formalizacion dentro de la misma transaccion
Call sbAsientoFormalizacion(Operacion.Operacion)

''Si es cheque verificar el número de Cheque a Registrar en Tesoreria
''y si el la transferencia del cheque a tesoreria es automatica
'strSQL = "select PideCheque from catalogo where codigo = '" & Operacion.Codigo & "'"
'rs.Open strSQL, glogon.Conection, adOpenStatic
'If Trim(UCase(cboTipoDocumento.Text)) = "CHEQUE" And rs!pideCheque = "S" And Len(vMensaje) = 0 Then
'   frmCR_SeguimientoDoc.Show vbModal
'   If Not Operacion.Valida Then vMensaje = vMensaje & vbCrLf & " - Se registró inconsistencias en la especificación del número del documento a desembolsar."
'End If 'CK
'rs.Close



Select Case fxTipoDocumento(cboTipoDocumento)
  Case "CK"
    'Crear Documento en Tesoreria Como Generado
    strSQL = "select PideCheque from catalogo where codigo = '" & Operacion.Codigo & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    If rs!pideCheque = "S" Then
       Call sbEnvioTesoreria(True, Operacion.Operacion)
    End If 'CK
    rs.Close
  Case "TE"
    'Enviar a Tesoreria como Pendiente
    Call sbEnvioTesoreria(False, Operacion.Operacion)
  Case Else
   'NADA
End Select

'BITACORA - FORMALIZACION
Call Bitacora("Registra", "Formalización de la OP: " & Operacion.Operacion)

'Cierra Transacciones
glogon.Conection.CommitTrans

Call sbCargaOperacion

MsgBox "Operación Modificada Satisfactoriamente ...", vbInformation

'Hacer Nota de Debito de Formalizacion
' Call sbDocumentoFormalizacion(Operacion.Operacion) 'Se Reemplaza por Cola de Asientos

Exit Sub

vError:
  Me.MousePointer = vbDefault
  glogon.Conection.RollbackTrans
  MsgBox Err.Description, vbCritical

  
End Sub

Private Sub cmdAplicaResolucion_Click()
Dim strSQL As String, strEstado As String, rs As New ADODB.Recordset
Dim vTramiteRapido As Integer


cmdAplicaResolucion.Enabled = False

If optResolucion(0).Value Then
  strEstado = "'A'"
Else
  strEstado = "'D'"
End If

rs.CursorLocation = adUseServer
rs.Open "select coalesce(count(*),0) as Existe from catalogo where tramite = 'R' and codigo ='" & Operacion.Codigo & "'", glogon.Conection, adOpenStatic
vTramiteRapido = rs!existe
rs.Close

'Verificar fechas
If fxVerificaResolucion(vTramiteRapido) Then


 If cboGarantia.Text = "Fiadores" Then
  
  If fxVerificaFiadores Then
    strSQL = "update reg_creditos set montoapr =" & CCur(medMontoAprobado) _
           & ", plazo= " & CInt(txtPlazoAprobado) & ",int = " & txtInteresAprobado _
           & ", interesv=" & txtInteresAprobado & ", cuota = " & CCur(txtCuotaAprobada) _
           & ", fechares = '" & Format(dtpFechaRes.Value, "yyyy/mm/dd") & "'" _
           & ", fechaforp = '" & Format(dtpFechaFormalizacion.Value, "yyyy/mm/dd") & "'" _
           & ", fecha_inicio_calculo ='" & Format(dtpDesembolso.Value, "yyyy/mm/dd") & "'" _
           & ", categoria_persona ='" & fxCalificacionPersona(Operacion.Cedula) & "'" _
           & ", primer_cuota = " & IIf((chkPrimera.Value = 1), "'S'", "'N'") _
           & ", estadosol = " & strEstado _
           & ", userres ='" & glogon.Usuario & "'" _
           & ", cuota_poliza = " & fxCuotaPolizaVida(CCur(medMontoAprobado), txtCodigo) _
           & ", cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex) _
           & ", cta_banco = '" & IIf(Len(Trim(txtCuentaAhorros)) = 0, 0, Trim(txtCuentaAhorros)) & "'" _
           & ", ind_deposito=" & IIf(UCase(cboTipoDocumento.Text) = "TRANSFERENCIA", 1, 0) _
           & ", emitir='" & fxTipoDocumento(cboTipoDocumento.Text) & "'" _
           & " where id_solicitud = " & Operacion.Operacion
    glogon.Conection.Execute strSQL
  Else
    MsgBox vMensaje, vbCritical
    Exit Sub
  End If 'Fiadores
 
 Else 'CboGarantia
 
    strSQL = "update reg_creditos set montoapr =" & CCur(medMontoAprobado) _
           & ", plazo= " & CInt(txtPlazoAprobado) & ",int = " & txtInteresAprobado _
           & ", interesv=" & txtInteresAprobado & ", cuota = " & CCur(txtCuotaAprobada) _
           & ", fechares = '" & Format(dtpFechaRes.Value, "yyyy/mm/dd") & "'" _
           & ", fechaforp = '" & Format(dtpFechaFormalizacion.Value, "yyyy/mm/dd") & "'" _
           & ", fecha_inicio_calculo ='" & Format(dtpDesembolso.Value, "yyyy/mm/dd") & "'" _
           & ", primer_cuota = " & IIf((chkPrimera.Value = 1), "'S'", "'N'") _
           & ", estadosol = " & strEstado _
           & ", userres ='" & glogon.Usuario & "'" _
           & ", categoria_persona ='" & fxCalificacionPersona(Operacion.Cedula) & "'" _
           & ", cuota_poliza = 0" _
           & ", cod_banco = " & cboBanco.ItemData(cboBanco.ListIndex) _
           & ", cta_banco = '" & IIf(Len(Trim(txtCuentaAhorros)) = 0, 0, Trim(txtCuentaAhorros)) & "'" _
           & ", ind_deposito=" & IIf(UCase(cboTipoDocumento.Text) = "TRANSFERENCIA", 1, 0) _
           & ", emitir='" & fxTipoDocumento(cboTipoDocumento.Text) & "'" _
           & " where id_solicitud = " & Operacion.Operacion
    glogon.Conection.Execute strSQL
 End If 'cboGarantia

 'Verificar si es un préstamo rapido y formalizarlo
 If strEstado = "'A'" Then
   
   'SUPUESTO: Ningun credito fiduciario deberia de ser con tramite rapido
   ' por lo de la poliza de vida
   If vTramiteRapido > 0 Then
    'Verificar Congelamiento
    If fxgCongelamiento(txtCedula, "per_cierra_AcCreditos") Then
      MsgBox "Esta Persona se encuentra CONGELADA, verifique...", vbExclamation
      Exit Sub
    End If
    'Formaliza Prestamo Rapido
    Call sbTramiteRapido(chkPrimera.Value)
   
   Else
    Call Bitacora("Registra", "Resolucion Aprobada a la OP: " & Operacion.Operacion)
    Call sbCargaOperacion
    MsgBox "Operación Modificada Satisfactoriamente ...", vbInformation
  End If
 
   
 Else
 
   Call Bitacora("Registra", "Resolucion Denegada a la OP: " & Operacion.Operacion)
   Call sbCargaOperacion
 
  
 End If 'strEstado
 

Else 'Verificacion Resolucion

 MsgBox vMensaje, vbCritical

End If


cmdAplicaResolucion.Enabled = True

End Sub

Private Sub sbAbonaRefundiciones(vOperacion As Long, strCodigo As String _
           , curMonto As Currency, curIntCor As Currency, curIntMor As Currency, vFecha As Date)
Dim rs As New ADODB.Recordset, strSQL As String, curTmpSaldo As Currency

If (curIntCor + curIntMor) > 0 Then 'Se Supone Que Cancela Toda la Mora
  rs.Open "select amortiza from vista_morosidad where id_solicitud =" & vOperacion, glogon.Conection, adOpenStatic
  If rs.EOF And rs.BOF Then
      curTmpSaldo = 0
  Else
      curTmpSaldo = rs!Amortiza
  End If
  rs.Close
  
  strSQL = "select id_solicitud,codigo,saldo from reg_creditos where id_solicitud =" & vOperacion
  rs.Open strSQL, glogon.Conection, adOpenStatic
   If Abs(rs!Saldo - curMonto) < 1 Then
     strSQL = "update reg_creditos set amortiza = amortiza + saldo,saldo=0,estado='C'" _
            & ",interesc = interesc +" & curIntCor + curIntMor _
            & " where id_solicitud = " & vOperacion
     glogon.Conection.Execute strSQL
   Else
     strSQL = "update reg_creditos set amortiza = amortiza + " & curMonto _
            & ",saldo=saldo-" & curMonto _
            & ",interesc=interesc+" & curIntCor + curIntMor _
            & " where id_solicitud =" & vOperacion
     glogon.Conection.Execute strSQL
   
   End If
   
   strSQL = "update morosidad set abintc=intc,abintm=intm,abamortiza=amortiza" _
          & ",fecult='" & Format(vFecha, "yyyy/mm/dd") & "',estado = 'C'" _
          & ",tcon = 3,ncon=" & Operacion.Operacion _
          & " where estado ='A' and id_solicitud=" & vOperacion
   glogon.Conection.Execute strSQL
   
   If curMonto - curTmpSaldo > 0 Then 'La Diferencia
        strSQL = "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
              & "FECHAP,TCON,NCON,ESTADO,ESTADO_ASIENTO) values('" & rs!Codigo & "'," _
              & rs!id_solicitud & ",0," & curMonto - curTmpSaldo _
              & ",0," & curMonto & ",'" & Format(vFecha, "yyyy/mm/dd") _
              & "'," & GLOBALES.glngFechaCR & ",3," & Operacion.Operacion & ",'A','G')"
        glogon.Conection.Execute strSQL
   End If
  rs.Close

Else '(curIntCor + curIntMor) > 0, No Hay Morosidad
  
  strSQL = "select id_solicitud,codigo,saldo from reg_creditos where id_solicitud =" & vOperacion
  rs.Open strSQL, glogon.Conection, adOpenStatic
   If Abs(rs!Saldo - curMonto) < 1 Then
     strSQL = "update reg_creditos set amortiza = amortiza + saldo,saldo=0,estado='C'" _
            & " where id_solicitud =" & vOperacion
     glogon.Conection.Execute strSQL
   Else
     strSQL = "update reg_creditos set amortiza = amortiza + " & curMonto _
            & ",saldo=saldo-" & curMonto _
            & " where id_solicitud =" & vOperacion
     glogon.Conection.Execute strSQL
   End If
   
   strSQL = "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
         & "FECHAP,TCON,NCON,ESTADO,ESTADO_ASIENTO) values('" & rs!Codigo & "'," _
         & rs!id_solicitud & ",0," & curMonto _
         & ",0," & curMonto & ",'" & Format(vFecha, "yyyy/mm/dd") _
         & "'," & GLOBALES.glngFechaCR & ",3," & Operacion.Operacion & ",'A','G')"
   glogon.Conection.Execute strSQL
  rs.Close

End If '(curIntCor + curIntMor) > 0

End Sub


Private Sub sbAbonaRetenciones(vOperacion As Long, strCodigo As String _
           , curMonto As Currency, curMora As Currency, vFecha As Date)
Dim rs As New ADODB.Recordset, strSQL As String, curTmpSaldo As Currency

strSQL = "select id_solicitud,codigo,plazo,cuota,amortiza" _
       & " from reg_creditos where id_solicitud =" & vOperacion
rs.Open strSQL, glogon.Conection, adOpenStatic
 
 If Abs((rs!Plazo * rs!cuota) - (curMonto + curMora + rs!Amortiza)) < 1 Then
   strSQL = "update reg_creditos set amortiza = amortiza + " & curMonto + curMora & ",estado='C'" _
          & " where id_solicitud =" & vOperacion
   glogon.Conection.Execute strSQL
 Else
   strSQL = "update reg_creditos set amortiza = amortiza + " & curMonto + curMora _
          & " where id_solicitud =" & vOperacion
   glogon.Conection.Execute strSQL
 End If
 
 If curMonto > 0 Then
    strSQL = "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
          & "FECHAP,TCON,NCON,ESTADO,ESTADO_ASIENTO) values('" & rs!Codigo & "'," _
          & rs!id_solicitud & ",0," & curMonto _
          & ",0," & curMonto & ",'" & Format(vFecha, "yyyy/mm/dd") _
          & "'," & GLOBALES.glngFechaCR & ",3," & Operacion.Operacion & ",'A','G')"
    glogon.Conection.Execute strSQL
 End If
 
 If curMora > 0 Then
    strSQL = "update morosidad set abintc = intc,abintm = intm,abamortiza = amortiza" _
          & ",fecult = '" & Format(vFecha, "yyyy/mm/dd") _
          & "',tcon = 3,ncon=" & Operacion.Operacion _
          & ",estado = 'C'" _
          & " where estado ='A' and id_solicitud=" & vOperacion
   glogon.Conection.Execute strSQL
 End If
rs.Close

End Sub


 

Private Sub sbFormalizar()
Dim rs As New ADODB.Recordset, strSQL As String
Dim curRefunde As Currency, curRetencion As Currency, curDesembolsos As Currency
Dim curIntDias As Currency, curInteres As Currency, curAmortiza As Currency
Dim lngPrideduc As Long, FechaUltima As Long, curPSD As Currency
Dim curCuota As Currency, vFecha As Date, curTotal As Currency
Dim curCargos As Currency, iMes As Integer, lngAnio As Long
Dim vFechaCalculo As Date


Me.MousePointer = vbHourglass

curRefunde = 0
curRetencion = 0
curIntDias = 0
curDesembolsos = 0
curInteres = 0
curAmortiza = 0
curPSD = 0
curCuota = 0
curCargos = 0


vFecha = fxFechaServidor

strSQL = "update reg_creditos set fechaforp = '" & Format(vFecha, "yyyy/mm/dd") _
       & "',fecha_inicio_calculo = '" & Format(dtpDesembolso.Value, "yyyy/mm/dd") _
       & "',cod_grupo = '" & fxCodigoCbo(cboRecursos) _
       & "' where id_solicitud = " & Operacion.Operacion
glogon.Conection.Execute strSQL

vDocumentoFormalizacion = False
vPasaFormalizacion = True

'curRefunde = fxMontoEnRefundiciones(Operacion.Operacion)
'curRetencion = fxMontoEnRetenciones(Operacion.Operacion)
'curDesembolsos = fxMontoEnDesembolsos(Operacion.Operacion)
'curCargos = fxMontoEnCargos(Operacion.Operacion)

'Se actualiza por Codigo Optimizado el 2003/02/19 por linea siguiente
curDesembolsos = fxMontoEnGeneral(Operacion.Operacion)


curIntDias = fxInteresesHastaFormalizar 'Ojo Con los Convenios

lngPrideduc = txtAno.Text & Format(fxConvierteMES(cboMes.Text), "00")
vFechaCalculo = fxFechaCalculo(Operacion.Codigo)


'Abonar Intereses + Primer Cuota

strSQL = "select Id_Solicitud,Codigo,FechaForp,Montoapr,Int,InteresV,Plazo" _
       & ",Cuota,Primer_Cuota,cuota_poliza,fecha_inicio_calculo" _
       & " from reg_creditos where id_solicitud =" & Operacion.Operacion
rs.Open strSQL, glogon.Conection, adOpenStatic


On Error GoTo vError

curPSD = IIf(IsNull(rs!cuota_poliza), 0, rs!cuota_poliza)


'Si la Primer Cuota esta Marcada, entonces :
' 1. Si la Fecha de Formalizacion es el 15 o antes, solo se le cobra la cuota y
'    se le eliminan los intereses de formalización, quedando solo los de la cuota
'    y la fecha de primer deduccion es el mes actual + 1
' 2. Si es despues del 15, se le cobran los intereses de formalización del dia hasta
'    el ultimo dia del mes en proceso + la primer cuota y la fecha de la primer
'    deducción es el mes de proceso + 2
' 3. Nota: La fecha de Calculo para el caso 1, tiene que ser un dia menor a la fecha
'    de formalizacion (para Reflejar efecto en la Boleta)
'                                               [*** Modificación al 2002/07/01 ***]

If rs!PRIMER_CUOTA = "S" Then
  curInteres = CCur(Format(rs!montoApr * rs!interesv / 1200, "################0.00"))
  curAmortiza = rs!cuota - curInteres
  curCuota = rs!cuota
  
  'Fecha de Ultima Deducción
  iMes = Month(rs!fecha_inicio_calculo)
  lngAnio = Year(rs!fecha_inicio_calculo)
  
'  If Day(rs!fechaforp) <= 15 Then
'     curIntDias = 0
'     vFechaCalculo = DateAdd("d", -1, rs!fechaforp)
'  Else
     
     If iMes = 12 Then
        iMes = 1
        lngAnio = lngAnio + 1
     Else
        iMes = iMes + 1
     End If
     
     'Calcular Intereses Hasta el Ultimo día del Mes
     vFechaCalculo = CDate(lngAnio & "/" & Format(iMes, "00") & "/01")
     vFechaCalculo = DateAdd("d", -1, vFechaCalculo)
     
     If curIntDias > 0 Then 'Esto porque los convenios no cobran intereses
       curIntDias = ((rs!interesv / 36000) * rs!montoApr * (Abs(DateDiff("d", vFechaCalculo, rs!fecha_inicio_calculo)) + 1))
     End If
     curIntDias = CCur(Format(curIntDias, "###############0.00"))
'  End If ' <= 15
  
  FechaUltima = CLng((lngAnio & Format(iMes, "00")))
  'Fin del Calculo de la Ultima Deducción

Else

  FechaUltima = fxFechaProcesoAnterior

End If


'Verificar Que el Monto Aprobado de la operacion sea mayor a las deducciones que
'Se le van ha aplicar.

curTotal = rs!montoApr - (curInteres + curAmortiza + curIntDias + curRefunde + curRetencion + curDesembolsos + curPSD + curCargos)


'Inicia Transacciones

glogon.Conection.BeginTrans


If curTotal < 0 Then
    If Abs(curTotal) > 0.009 Then
       vPasaFormalizacion = False
       glogon.Conection.RollbackTrans
       Me.MousePointer = vbDefault
       MsgBox " - No se puede formalizar esta operación porque el Monto de Deducciones es Mayor al Monto Aprobado...", vbCritical
       Exit Sub
    End If
End If
'Verifica si el monto a girar es Cero, para el cual se debe Generar Forzadamente la ND
If (curTotal < 0 And curTotal > -0.009) Or curTotal = 0 Then
' If rs!montoapr - (curInteres + curAmortiza + curIntDias + curRefunde + curDesembolsos + curPSD) Then
  vDocumentoFormalizacion = True
End If

  'No se incluye PSD porque será referenciada en otro procedimiento
  strSQL = "insert creditos_dt(CODIGO,ID_SOLICITUD,CUOTA,ABONO,INTCP,AMORTIZA,FECHAS," _
         & "FECHAP,TCON,NCON,ESTADO,ESTADO_ASIENTO) values('" & Operacion.Codigo & "'," _
         & Operacion.Operacion & "," & curCuota & "," & curInteres + curIntDias + curAmortiza _
         & "," & curInteres + curIntDias & "," & curAmortiza & ",'" & Format(vFecha, "yyyy/mm/dd") _
         & "'," & GLOBALES.glngFechaCR & ",2," & Operacion.Operacion & ",'A','G')"
  glogon.Conection.Execute strSQL
rs.Close

'Rebajar PSD Tambien al monto a girar
strSQL = "update reg_creditos set pagare = " & txtPagare _
       & ",documento_referido='" & Mid(txtDocumento, 1, 18) & "', Estadosol = 'F',Estado ='A'" _
       & ",prideduc =" & lngPrideduc & ",monto_girado = " _
       & Operacion.MontoAprobado - (curRefunde + curRetencion + curDesembolsos + curInteres + curIntDias + curAmortiza + curPSD + curCargos) _
       & ",fecult = " & FechaUltima _
       & ",fecha_calculo_int = '" & Format(vFechaCalculo, "yyyy/mm/dd") & "'" _
       & ",userfor = '" & glogon.Usuario & "'" _
       & ",saldo_mes = montoapr - " & curAmortiza _
       & ",saldo = montoapr - " & curAmortiza & ", amortiza = " & curAmortiza & ", interesc = " & curInteres + curIntDias
         
  If (chkEnviarATesoreria.Value = 0) Then
   strSQL = strSQL & ",tesoreria='" & Format(vFecha, "yyyy/mm/dd") & "'"
   vDocumentoFormalizacion = True
  End If 'CON SOLO QUE NO SE ENVIA A TESORERIA HAY QUE CREAR NOTA DE DEBITO
  
  If (chkEnviarATesoreria.Value = 1) And vDocumentoFormalizacion And fxMontoEnDesembolsos(Operacion.Operacion) = 0 Then
   strSQL = strSQL & ",tesoreria = '" & Format(vFecha, "yyyy/mm/dd") & "'"
   vDocumentoFormalizacion = True
  End If 'ES ND Y NO NECESITA ENVIARSE A TESORERIA, PORQUE NO TIENE DESEMBOLSOS
         'DE LO CONTRATIO SE GENERA DE LA ND Y SE TRASLADAN LOS DESEMBOLSOS A TESORERIA
         'EN OTRO PROCESO.
  
  
  strSQL = strSQL & " where id_solicitud = " & Operacion.Operacion
  glogon.Conection.Execute strSQL

'Abonar Refundiciones
strSQL = "select id_solicitud,codigo,Monto,IntCor,IntMor from refundiciones where id_solicitudr = " & Operacion.Operacion
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  Call sbAbonaRefundiciones(rs!id_solicitud, rs!Codigo, rs!Monto, rs!IntCor, rs!IntMor, vFecha)
  rs.MoveNext
Loop
rs.Close

'Abonar Retenciones
strSQL = "select id_solicitud,codigo,monto,mora from refunde_retencion where id_solicitudr = " & Operacion.Operacion
rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  Call sbAbonaRetenciones(rs!id_solicitud, rs!Codigo, rs!Monto, rs!mora, vFecha)
  rs.MoveNext
Loop
rs.Close

'Crea Asiento de Formalizacion
Call sbAsientoFormalizacion(Operacion.Operacion)

Select Case fxTipoDocumento(cboTipoDocumento)
  Case "CK"
    'Crear Documento en Tesoreria Como Generado
    strSQL = "select PideCheque from catalogo where codigo = '" & Operacion.Codigo & "'"
    rs.Open strSQL, glogon.Conection, adOpenStatic
    If rs!pideCheque = "S" Then
       strSQL = "update reg_creditos set documento_referido = '" & fxTipoDocumento(cboTipoDocumento) _
              & "-" & Operacion.Documento & "',tesoreria='" & Format(vFecha, "yyyy/mm/dd") _
              & "' where id_solicitud = " & Operacion.Operacion
       glogon.Conection.Execute strSQL
       Call sbEnvioTesoreria(True, Operacion.Operacion)
    End If 'CK
    rs.Close
  Case Else
   'NADA
End Select


'BITACORA
Call Bitacora("Registra", "Formalización de la OP: " & Operacion.Operacion)

'Cierra Transacciones
glogon.Conection.CommitTrans

Me.MousePointer = vbDefault

MsgBox "Formalización Aplicada Satisfactoriamente...", vbInformation

Exit Sub

vError:
  glogon.Conection.RollbackTrans
  Me.MousePointer = vbDefault
  MsgBox Err.Description, vbCritical
  
End Sub

Private Sub sbAnular()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vProcedimiento As Integer, vMensaje As String
Dim curInteres As Currency, curAmortiza As Currency
Dim rs2 As New ADODB.Recordset, vND As Currency

Me.MousePointer = vbHourglass

vMensaje = ""

'Indica si es una nota de debito para lo cual el sistema debe de generar
'Una nota de credito para reversar el movimiento contable
vND = 0

strSQL = "select * from reg_creditos where id_solicitud =" & Operacion.Operacion
rs.Open strSQL, glogon.Conection, adOpenStatic

If rs!fecult >= rs!prideduc Then
  rs.Close
  MsgBox "La Anulación no procede, porque ya se realizarón Deducciones de Planilla", vbCritical
  Me.MousePointer = vbDefault
  Exit Sub
End If

If Format(rs!fechaforp, "yyyy/mm/dd") < Format(fxFechaServidor, "yyyy/mm/dd") Then
  rs.Close
  MsgBox "La Anulación no procede, porque la formalizacion se realizó otro día", vbCritical
  Me.MousePointer = vbDefault
  Exit Sub
End If


vND = Val(rs!ndocumento)

If Not IsNull(rs!TESORERIA) Then
    
    vMensaje = "- ESTA OPERACION FUE ENVIADA A TESORERIA" _
             & ", DEBE DE INDICAR A LOS TESOREROS QUE NO EMITAN EL DOCUMENTO RESPECTIVO O" _
             & " QUE ANULEN LA EMISION." & vbCrLf & vbCrLf & " LA SOLICITUD SE ENVIO EL : " _
             & rs!TESORERIA
End If

On Error GoTo vError
'Inicia Transacciones
glogon.Conection.BeginTrans

If rs!estadosol = "F" Then
    strSQL = "update reg_creditos set estadosol ='N',estado='N' where id_solicitud=" & Operacion.Operacion
    glogon.Conection.Execute strSQL
    vProcedimiento = 1
Else
    strSQL = "update reg_creditos set estadosol ='N' where id_solicitud=" & Operacion.Operacion
    glogon.Conection.Execute strSQL
    vProcedimiento = 0
End If
rs.Close

If vProcedimiento = 1 Then
 'Anular Abonos a Morosidad y Registros En Detalle de las refundiciones
 'Finalizar con NC
 strSQL = "select * from refundiciones where id_solicitudr =" & Operacion.Operacion
 rs.Open strSQL, glogon.Conection, adOpenStatic
 Do While Not rs.EOF
  strSQL = "update reg_creditos set estado = 'A',saldo=saldo + " & rs!Monto _
         & ", amortiza=amortiza-" & rs!Monto & ",interesc=interesc-" & rs!IntCor + rs!IntMor _
         & " where id_solicitud=" & rs!id_solicitud
  glogon.Conection.Execute strSQL
  rs.MoveNext
 Loop
 rs.Close
 
'Anula las retenciones refundidas
 strSQL = "select * from refunde_retencion where id_solicitudr =" & Operacion.Operacion
 rs.Open strSQL, glogon.Conection, adOpenStatic
 Do While Not rs.EOF
  strSQL = "update reg_creditos set estado = 'A'" _
         & ", amortiza=amortiza-" & (rs!Monto + rs!mora) _
         & " where id_solicitud=" & rs!id_solicitud
  glogon.Conection.Execute strSQL
  rs.MoveNext
 Loop
 rs.Close

 'Elimina los registros de refundicion generados por la formalizacion
 'tanto en refundiciones de credito como de retenciones
 
  strSQL = "delete creditos_dt where tcon=3 and ncon=" & Operacion.Operacion
  glogon.Conection.Execute strSQL
  
  strSQL = "update morosidad SET estado = 'A',abintc = 0,abintm = 0,abamortiza =0 where tcon=3 and ncon=" & Operacion.Operacion
  glogon.Conection.Execute strSQL
 
 'Anular Registros Creditos DT de la operacion
  strSQL = "delete creditos_dt where id_solicitud=" & Operacion.Operacion
  glogon.Conection.Execute strSQL

  'Asiento de Reversion
  Call sbAsientoAnulacionFormaliza(Operacion.Operacion)

End If 'Procedimiento


'BITACORA
Call Bitacora("Registra", "Anulación de la OP: " & Operacion.Operacion)

'Cierra Transacciones
glogon.Conection.CommitTrans

'''Crear Nota de Credito con la Reversion de la nota de debito
''' ************ ACTIVAR ESTO PARA UTILIZAR PROCEDIMIENTOS DE NOTAS DE DEBITO *************
''If vProcedimiento = 1 And vND > 0 Then
'' Call sbDocumentoAnulacionFormalizacion(Operacion.Operacion)
''End If

vMensaje = "Anulación Realizada Satisfactoriamente..."

Me.MousePointer = vbDefault

If Len(vMensaje) > 0 Then MsgBox vMensaje, vbInformation


Exit Sub

vError:
 glogon.Conection.RollbackTrans
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical
End Sub

Private Sub cmdAplicarFormalizacion_Click()

If Me.optFormalizacion(0).Value Then
  'Verificar Congelamiento
  If fxgCongelamiento(txtCedula, "per_cierra_AcCreditos") Then
    MsgBox "Esta Persona se encuentra CONGELADA, verifique...", vbExclamation
    Exit Sub
  End If
 
 If fxVerificaFormalizacion Then
    Call sbFormalizar
 Else 'Falla Verificacion de Formalizacion
  MsgBox vMensaje, vbCritical
 End If

Else 'Anulacion de la formalizacion
  If fxVerificaAnulacion Then
      Call sbAnular
  Else
    MsgBox vMensaje, vbCritical
  End If
End If

Call sbCargaOperacion

End Sub

Private Sub cmdCalculoOperacion_Click()
frmCR_CalculoOperacion.Show vbModal
End Sub

Private Sub cmdCargos_Click()
On Error GoTo vError
If Operacion.Operacion > 0 Then
 Operacion.Ventana = "C"
 frmCR_SeguimientoReqCar.Show vbModal
End If

Exit Sub
vError:
End Sub

Private Sub cmdDatosPersonales_Click()
If Operacion.Operacion > 0 Then
 GLOBALES.gCedulaActual = Operacion.Cedula
 frmCR_VerificaDatosPersonales.Show vbModal
End If
End Sub

Private Sub cmdDeduccionPlanilla_Click()
Dim strSQL As String

If chkDeducirPlanilla.Value = 1 Then
 strSQL = "update reg_creditos set IND_DEDUCE_PLANILLA='S' where " _
    & "id_solicitud=" & Operacion.Operacion
  Call Bitacora("Registra", "Indica la Deducción de Planilla de la OP: " & Operacion.Operacion)
Else
strSQL = "update reg_creditos set IND_DEDUCE_PLANILLA='N' where " _
    & "id_solicitud=" & Operacion.Operacion
  Call Bitacora("Registra", "Indica la NO Deducción de Planilla de la OP: " & Operacion.Operacion)
End If
glogon.Conection.Execute strSQL

MsgBox "Actualización Realizada...", vbInformation

End Sub

Private Sub cmdFiadores_Click()
On Error GoTo vError
If Operacion.Operacion > 0 Then frmCR_SolicitudesFiadores.Show vbModal
Exit Sub
vError:
frmCR_SolicitudesFiadores.SetFocus
 
End Sub

Private Sub cmdFirmas_Click()
If Operacion.Operacion > 0 Then frmCR_SeguimientoFirmas.Show vbModal
End Sub


Private Sub cmdRequisitos_Click()
On Error GoTo vError
If Operacion.Operacion > 0 Then
 Operacion.Ventana = "R"
 frmCR_SeguimientoReqCar.Show vbModal
End If
Exit Sub
vError:
End Sub



Private Sub dtpFechaFormalizacion_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
If KeyCode = vbKeyReturn Then chkPrimera.SetFocus
End Sub

Private Sub dtpFechaFormalizacion_Change()
Dim strSQL As String

strSQL = "update reg_creditos set fechaforp = '" & Format(dtpFechaFormalizacion.Value, "yyyy/mm/dd") & "'" _
       & " where id_solicitud = " & Operacion.Operacion
    
glogon.Conection.Execute strSQL

Call Bitacora("Modifica", "Fecha Formalizacion Operacion " & Operacion.Operacion)

End Sub

Private Sub dtpFechaRes_Change()
Dim strSQL As String

strSQL = "update reg_creditos set fechares = '" & Format(dtpFechaRes.Value, "yyyy/mm/dd") & "'" _
       & " where id_solicitud = " & Operacion.Operacion
    
glogon.Conection.Execute strSQL

Call Bitacora("Modifica", "Fecha Resolución Operacion " & Operacion.Operacion)

End Sub

Private Sub dtpFechaRes_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then dtpFechaFormalizacion.SetFocus
End Sub

Private Sub dtpFechaSolicitud_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyReturn Then cboEstado.SetFocus
End Sub


Private Function fxVerificaExisteCodigo(strCodigo As String) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
strSQL = "select coalesce(count(*),0) as Existe from catalogo where codigo ='" & strCodigo & "'"
rsX.Open strSQL, glogon.Conection, adOpenStatic
fxVerificaExisteCodigo = IIf((rsX!existe > 0), True, False)
rsX.Close
End Function

Private Function fxVerificaExisteRangoCodigo(strCodigo As String, curMonto As Currency) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
 
strSQL = "select coalesce(count(*),0) as Existe from rangos"
strSQL = strSQL & " where codigo ='" & strCodigo & "' and " & curMonto & " >=de and " _
        & curMonto & " <=  hasta"
rsX.Open strSQL, glogon.Conection, adOpenStatic
fxVerificaExisteRangoCodigo = IIf((rsX!existe > 0), True, False)
rsX.Close
End Function

Private Function fxVerificaFiadores() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
fxVerificaFiadores = True
vMensaje = ""
strSQL = "select coalesce(count(*),0) as Existe from fiadores where estado ='A' and id_solicitud=" & Operacion.Operacion
rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX!existe = 0 Then vMensaje = "- La garantía de esta operacion es fiduciaria y no se ha especificado ningún fiador"
rsX.Close

If Len(vMensaje) > 0 Then fxVerificaFiadores = False

End Function

Private Function fxVerificaFirmas() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String

fxVerificaFirmas = True
vMensaje = ""
strSQL = "select firma from fiadores where estado ='A' and id_solicitud=" & Operacion.Operacion
rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic
Do While rsX.EOF = False
  If rsX!firma = "N" Then
    vMensaje = "- El No se han registrado todas las firmas de los fiadores"
  End If
  rsX.MoveNext
Loop
rsX.Close

strSQL = "select firma_deudor from reg_creditos where id_solicitud=" & Operacion.Operacion
rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic
 If IIf(IsNull(rsX!firma_deudor), 0, rsX!firma_deudor) = 0 Then vMensaje = vMensaje & vbCrLf & "- No se ha registrado la firma del deudor"
rsX.Close
If Len(vMensaje) > 0 Then fxVerificaFirmas = False

End Function

Private Function fxVerificaRecepcion() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String, vEstadoPersona As String
Dim vEstado As String

fxVerificaRecepcion = True
vMensaje = ""

strSQL = "select estadoactual from socios where cedula ='" & Trim(txtCedula) & "'"
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX.EOF And rsX.BOF Then
  vEstado = ""
Else
  vEstado = rsX!estadoactual & ""
End If
rsX.Close


If IsNumeric(txtPlazoSolicitado) Then
 If txtPlazoSolicitado < 1 Then vMensaje = vMensaje & vbCrLf & "- El Plazo Solicitado es Inválido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Plazo Solicitado es Inválido"
End If
If IsNumeric(txtInteresSolicitado) Then
 If txtInteresSolicitado < 0 Or txtInteresSolicitado > 100 Then vMensaje = vMensaje & vbCrLf & "- El Interés Solicitado es Inválido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Interés Solicitado es Inválido"
End If

If IsNumeric(medMontoSolicitado.Text) Then
 If medMontoSolicitado.Text < 1 Then vMensaje = vMensaje & vbCrLf & "- El Monto Solicitado Solicitado es Inválido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Monto Solicitado es Inválido"
End If

'Verifica que el Banco de Deposito exista o este asignado (autorizado para el usuario)
If fxEstadoOperacion(cboEstado.Text) = "P" Or fxEstadoOperacion(cboEstado.Text) = "R" Then
    If Not fxBancoAsignado(cboBanco.ItemData(cboBanco.ListIndex), glogon.Usuario) Then
       vMensaje = vMensaje & vbCrLf & "- EL BANCO INDICADO NO SE ENCUENTRA AUTORIZADO AL USUARIO : " & glogon.Usuario
    End If
End If

'VERIFICAR SI TIENE CODIFICACION CONTABLE

strSQL = "select * from catalogo where codigo ='" & txtCodigo & "'"
rsX.Open strSQL, glogon.Conection, adOpenStatic
 If rsX.EOF And rsX.BOF Then
   vMensaje = vMensaje & vbCrLf & "- El código de préstamo no existe"
 Else
   
  
  'Verifica si el codigo tiene codificacion contable
  'Es suficiente con evaluar cualquiera de las 9, pues el sistema
  'solo permite actualizar cuando se especifican todas.
   If IsNull(rsX!ctaNintC) Then vMensaje = vMensaje & vbCrLf & "- El código no se encuentra codificado contablemente"
   
   'No se permiten retenciones ni polizas
   If rsX!retencion = "S" Or rsX!poliza = "S" Then vMensaje = vMensaje & vbCrLf & "- No se permite guardar porque el código pertenece a una Retencion o Poliza"
  
   'Verificar que el Codigo se encuentre Activo
   If rsX!activo = 0 Then vMensaje = vMensaje & vbCrLf & "- La Línea de Crédito no se encuentra Activa..."
  
   'Verificar los Estados de la Persona Válidos para Esta Linea
   Select Case vEstado
     Case "S"
        If rsX!EstatusS = 0 Then vMensaje = vMensaje & vbCrLf & "- Esta Línea de Crédito no Admite Socios Activos..."
     Case "A"
        If rsX!EstatusA = 0 Then vMensaje = vMensaje & vbCrLf & "- Esta Línea de Crédito no Admite Ex-Socios Internos..."
     Case "P"
        If rsX!EstatusP = 0 Then vMensaje = vMensaje & vbCrLf & "- Esta Línea de Crédito no Admite Ex-Empleados (Patronal)..."
     Case "N"
        If rsX!EstatusN = 0 Then vMensaje = vMensaje & vbCrLf & "- Esta Línea de Crédito no Admite No Socios..."
     Case ""
       vMensaje = vMensaje & vbCrLf & "- La Persona no Existe"
   End Select
   
 End If
rsX.Close

'VERIFICAR COMBOS
If fxCodigoDestino(fxCodigoCbo(cboDestino), txtCodigo) = 0 Then vMensaje = vMensaje & vbCrLf & "- El Destino No es válido para Esta Línea"
If fxCodigoBanco(cboBanco.Text) = 0 Then vMensaje = vMensaje & vbCrLf & "- El Banco Especificado NO EXISTE"
If fxCodigoComite(cboComite.Text) = 0 Then vMensaje = vMensaje & vbCrLf & "- El Comité Especificado NO EXISTE"
If fxEstadoOperacion(cboEstado.Text) = "" Then vMensaje = vMensaje & vbCrLf & "- El Estado de la Operación NO ES VALIDO"
If fxTipoDocumento(cboTipoDocumento.Text) = "" Then vMensaje = vMensaje & vbCrLf & "- La emisión de la operación NO ES VALIDA"
If fxGarantia(cboGarantia.Text) = "" Then vMensaje = vMensaje & vbCrLf & "- La Garantía especificada NO ES VALIDA"


'Verificar que la persona no tenga prestamos en Cobro Judicial Activos
strSQL = "select coalesce(count(*),0) as Existe from reg_creditos" _
       & " where estado = 'A' and proceso = 'J' and cedula = '" & txtCedula & "'"
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX!existe > 0 Then vMensaje = vMensaje & vbCrLf & "- La persona tiene créditos en Cobro Judicial"
rsX.Close

If Len(vMensaje) > 0 Then fxVerificaRecepcion = False

End Function

Private Function fxVerificaNivel()
Dim rsX As New ADODB.Recordset, rsX2 As New ADODB.Recordset, strSQL As String

strSQL = "select count(*) as Existe from nivel_miembros A, nivel_derechos B where A.nv_cod_grupo = " _
       & "B.nv_cod_grupo and nombre = '" & glogon.Usuario & "' and codigo = '" _
       & txtCodigo & "'"
rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic

rsX.Close

End Function

Private Function fxVerificaResolucion(vTramiteRapido As Integer) As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String, iMes As Integer, lngAnio As Long
Dim Porcentaje As Currency, curMonto As Currency, vFecha As Date, vFechaCalculo As Date

vMensaje = ""

fxVerificaResolucion = True

vFecha = Format(fxFechaServidor, "yyyy/mm/dd")

If dtpFechaRes.Value > dtpFechaFormalizacion.Value Then
  dtpFechaFormalizacion.Value = vFecha
  strSQL = "update reg_creditos set fechaforp = '" & Format(dtpFechaFormalizacion.Value, "yyyy/mm/dd") _
         & "' where id_solicitud = " & Operacion.Operacion
  glogon.Conection.Execute strSQL
End If

If IsNumeric(medMontoAprobado) Then
 If medMontoAprobado.Text < 1 Then
   vMensaje = vMensaje & vbCrLf & "- El monto Aprobado es Inválido"
 Else
        strSQL = "select COALESCE(count(*),0) as Existe" _
           & " from NIVEL_GRUPOS N INNER JOIN nivel_miembros A" _
           & " ON N.NV_COD_GRUPO = A.NV_COD_GRUPO INNER JOIN nivel_derechos B" _
           & " ON N.NV_COD_GRUPO = B.NV_COD_GRUPO Where A.nombre = '" & glogon.Usuario _
           & "' and B.codigo = '" & txtCodigo & "' AND N.nv_tipo = 'R'" _
           & " and (" & CCur(medMontoAprobado.Text) & " between nv_desde and nv_hasta)"
    rsX.CursorLocation = adUseServer
    rsX.Open strSQL, glogon.Conection, adOpenStatic
    If rsX!existe = 0 Then
      vMensaje = vMensaje & vbCrLf & "- No existe nivel de resolución de este usuario para el código " & txtCodigo
    End If
    rsX.Close
 End If
Else
 vMensaje = vMensaje & vbCrLf & "- El monto Aprobado es Inválido"
End If



If IsNumeric(txtPlazoAprobado) Then
 If txtPlazoAprobado < 1 Then vMensaje = vMensaje & vbCrLf & "- El Plazo Aprobado es Inválido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Plazo Aprobado es Inválido"
End If

If IsNumeric(txtInteresAprobado) Then
 If txtInteresAprobado < 0 Then vMensaje = vMensaje & vbCrLf & "- El Interes Aprobado es Inválido"
Else
   vMensaje = vMensaje & vbCrLf & "- El Interes Aprobado es Inválido"
End If

If UCase(cboGarantia.Text) = "SOBRE AHORROS" Then
    strSQL = "select CR_POR_AHORRO  from par_ahcr"
    rsX.Open strSQL, glogon.Conection, adOpenStatic
       Porcentaje = rsX!Cr_Por_ahorro
    rsX.Close
    
    strSQL = "select coalesce(AHORRO,0) as Ahorro, coalesce(Capitaliza,0) as Capital" _
           & " from ahorro_consolidado where cedula = '" & txtCedula & "'"
    rsX.Open strSQL, glogon.Conection, adOpenStatic
    If CCur(medMontoAprobado.Text) > ((rsX!Ahorro + rsX!capital) * (Porcentaje / 100)) Then
       vMensaje = vMensaje & vbCrLf & " - El Monto aprobado excede el " & Porcentaje _
                & "% de sus ahorros"
    End If
    rsX.Close
End If

'Porcentaje de la Garantia sobre fondos
If UCase(cboGarantia.Text) = "FONDOS / PLANES" Then
 If CCur(medMontoAprobado.Text) > fxDisponibleFondos(txtCedula) Then
     vMensaje = vMensaje & vbCrLf & " - El Monto aprobado excede el " & Porcentaje _
              & "% de sus FONDOS"
 End If
End If

If vTramiteRapido > 0 Then
  'Verificar que las deducciones sean igual o mayor que el monto aprobado
  'Incluir Morosidad de Todas Las Operaciones en Refundiciones de la Persona
  curMonto = 0
  strSQL = "select R.saldo,Coalesce(V.Amortiza,0) as Amortiza " _
         & " from reg_creditos R left join Vista_Morosidad V on R.id_solicitud = V.id_solicitud" _
         & " where R.estado = 'A' and R.Saldo > 0" _
         & " and R.cedula = '" & Operacion.Cedula & "' and R.codigo = '" _
         & Operacion.Codigo & "' and R.id_solicitud < " & Operacion.Operacion
  rsX.CursorLocation = adUseServer
  rsX.Open strSQL, glogon.Conection, adOpenStatic
    If Not rsX.EOF And Not rsX.BOF Then curMonto = curMonto + rsX!Saldo - rsX!Amortiza
  rsX.Close
   
  'Sacar Mora de Todas las Operaciones
  strSQL = "select coalesce(sum(M.intc),0) + coalesce(sum(M.intm),0) + coalesce(sum(M.amortiza),0) as Mora" _
         & " from reg_creditos R inner join Morosidad M on R.id_solicitud = M.id_solicitud" _
         & " where R.estado = 'A' and R.Saldo > 0 and M.estado = 'A'" _
         & " and cedula = '" & Operacion.Cedula & "'"
  rsX.CursorLocation = adUseServer
  rsX.Open strSQL, glogon.Conection, adOpenStatic
    If Not rsX.EOF And Not rsX.BOF Then curMonto = curMonto + rsX!mora
  rsX.Close
   
  If medMontoAprobado.Text = "" Then medMontoAprobado.Text = 0
   

  
  If curMonto > CCur(medMontoAprobado.Text) Then vMensaje = vMensaje & vbCrLf & " - El Monto aprobado es Inferior a las deducciones que se le efectuaran"
  
    'Si es cheque verificar el número de Cheque a Registrar en Tesoreria
    'y si el la transferencia del cheque a tesoreria es automatica
    strSQL = "select PideCheque from catalogo where codigo = '" & Operacion.Codigo & "'"
    rsX.Open strSQL, glogon.Conection, adOpenStatic
    If Trim(UCase(cboTipoDocumento.Text)) = "CHEQUE" And rsX!pideCheque = "S" And Len(vMensaje) = 0 Then
       frmCR_SeguimientoDoc.Show vbModal
       If Not Operacion.Valida Then vMensaje = vMensaje & vbCrLf & " - Se registró inconsistencias en la especificación del número del documento a desembolsar."
    End If 'CK
    rsX.Close

  
  
End If

  If chkPrimera.Value = vbChecked Then
    curMonto = curMonto + CCur(txtCuotaAprobada)
'       If Day(vFecha) > 15 Then
            iMes = Month(vFecha)
            lngAnio = Year(vFecha)
            If iMes = 12 Then
               iMes = 1
               lngAnio = lngAnio + 1
            Else
               iMes = iMes + 1
            End If
            'Calcular Intereses Hasta el Ultimo día del Mes
            vFechaCalculo = CDate(lngAnio & "/" & Format(iMes, "00") & "/01")
            vFechaCalculo = DateAdd("d", -1, vFechaCalculo)
            
            If fxInteresesHastaFormalizar(dtpFechaFormalizacion.Value, CCur(medMontoAprobado.Text)) > 0 Then 'Esto porque los convenios no cobran intereses
              curMonto = curMonto + ((CCur(txtInteresAprobado) / 36000) * CCur(medMontoAprobado.Text) * (Abs(DateDiff("d", vFechaCalculo, vFecha)) + 1))
            End If
'       End If ' > 15
    
  Else
  
    'Calculo de Interes x Dias (Normalmente Utilizado)
    curMonto = curMonto + fxInteresesHastaFormalizar(dtpFechaFormalizacion.Value, CCur(medMontoAprobado.Text))
  
  End If


If Len(vMensaje) > 0 Then fxVerificaResolucion = False
'Verifica si no sobre pasa el porcentaje de sobre ahorros (Refundiciones)

If vTramiteRapido = 0 Then 'Si no es tramite rapido pregunta refundiciones
    If UCase(cboGarantia.Text) = "SOBRE AHORROS" _
       And fxVerificaResolucion And optResolucion(0).Value Then
       Operacion.Operacion = txtOperacion
       Operacion.Codigo = txtCodigo
       Operacion.MontoAprobado = medMontoAprobado.Text
       frmCR_SeguimientoSobreAhorros.Show vbModal
       fxVerificaResolucion = Operacion.Valida
       vMensaje = vMensaje & vbCrLf & "- Excede el porcentaje sobre ahorros..."
    End If
    
    'Verifica si la persona está o no bloqueda para formalizaciones
    strSQL = "select bloqueo from socios where cedula = '" & Operacion.Cedula & "'"
    rsX.Open strSQL, glogon.Conection, adOpenStatic
    If rsX!bloqueo = 1 Then
       vMensaje = vMensaje & vbCrLf & "- Esta persona se encuentra bloqueda, hasta mañana se le podran formalizar operaciones..."
    End If
    rsX.Close
End If


End Function

Private Function fxVerificaFormalizacion() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngPrideduc As Long, vFecha As Date
Dim Porcentaje As Double, vMontoRefunde As Currency

vMensaje = ""
fxVerificaFormalizacion = True

vFecha = fxFechaServidor

'Verificar Si la Linea de Credito este Activa


'Verificar Si el Destino Existe y esta Activo


If Not fxVerificaFirmas Then vMensaje = vMensaje & vbCrLf & "- No se han registrado todas las firmas..."

If dtpFechaFormalizacion.Value > dtpDesembolso.Value Then vMensaje = vMensaje & vbCrLf & "- La fecha del desembolsos no puede ser menor que la fecha de formalizacion"

If (Operacion.Estado = "A" Or Operacion.Estado = "C") And Me.optFormalizacion(0).Value = True _
    Then vMensaje = vMensaje & vbCrLf & "- Esta Operación ya fue procesada"

If Not fxVerificaCodigoDoble(Operacion.Codigo, Operacion.Cedula, Operacion.Operacion) _
   Then vMensaje = vMensaje & vbCrLf & "- No se permite sobrepasar el numero maximo de operaciones en esta linea"

If IsNumeric(txtPagare) Then
  If txtPagare < 0 Then vMensaje = vMensaje & vbCrLf & "- # de Pagaré no es válido"
Else
  vMensaje = vMensaje & vbCrLf & "- # de Pagaré no es válido"
End If

If IsNumeric(txtAno) Then
  If txtAno < Year(vFecha) Then vMensaje = vMensaje & vbCrLf & "- El Año especificado no es válido"
Else
  vMensaje = vMensaje & vbCrLf & "- El Año para la primer deduccion no es válido"
End If

If fxConvierteMES(cboMes.Text) = cboMes.Text Then vMensaje = vMensaje & vbCrLf & "- El Mes para la primer deduccion no es válido"

lngPrideduc = txtAno.Text & Format(fxConvierteMES(cboMes.Text), "00")

If lngPrideduc <= GLOBALES.glngFechaCR Then vMensaje = vMensaje & vbCrLf & "- La primer deducción no es válida porque es igual o menor a la fecha de proceso actual"

If Month(dtpFechaFormalizacion.Value) <> Month(vFecha) Or Year(dtpFechaFormalizacion.Value) <> Year(vFecha) Then
 'Actualiza la fecha de formalizacion
 strSQL = "update reg_creditos set fechaforp = '" & Format(vFecha, "yyyy/mm/dd") _
        & "' where id_solicitud = " & Operacion.Operacion
 glogon.Conection.Execute strSQL
 dtpFechaFormalizacion.Value = vFecha
End If

strSQL = "select COALESCE(count(*),0) as Existe" _
   & " from NIVEL_GRUPOS N INNER JOIN nivel_miembros A" _
   & " ON N.NV_COD_GRUPO = A.NV_COD_GRUPO INNER JOIN nivel_derechos B" _
   & " ON N.NV_COD_GRUPO = B.NV_COD_GRUPO Where A.nombre = '" & glogon.Usuario _
   & "' and B.codigo = '" & txtCodigo & "' AND N.nv_tipo = 'F'" _
   & " and (" & CCur(medMontoAprobado.Text) & " between nv_desde and nv_hasta)"

rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX!existe = 0 Then
  vMensaje = vMensaje & vbCrLf & "- No existe nivel de formalización de este usuario para el código " & txtCodigo
End If
rsX.Close


'Verifica que los saldos de las refundiciones ingresadas esten actualizados

strSQL = "select R.id_solicitud,R.codigo,R.monto,C.saldo" _
       & " from refundiciones R inner join reg_creditos C on R.id_solicitud = C.id_solicitud" _
       & " where R.id_solicitudr=" & Operacion.Operacion
rsX.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rsX.EOF
  If rsX!Monto > rsX!Saldo Then
      vMensaje = vMensaje & vbCrLf & "- El saldo a refundir vario en la operación : " & rsX!id_solicitud
  End If
  rsX.MoveNext
Loop
rsX.Close

'Se incluye para Verificar el 90% del ahorro
If UCase(cboGarantia.Text) = "SOBRE AHORROS" Then
    strSQL = "select CR_POR_AHORRO  from par_ahcr"
    rsX.Open strSQL, glogon.Conection, adOpenStatic
    Porcentaje = rsX!Cr_Por_ahorro
    rsX.Close
    
        
    strSQL = "select coalesce(sum(saldo),0) as Saldos from reg_creditos where estado = 'A'" _
           & " and saldo > 0 and garantia = 'A' and cedula = '" & Operacion.Cedula & "'"
    rsX.Open strSQL, glogon.Conection, adOpenStatic
    vMontoRefunde = rsX!saldos  'Saldos Sobre Ahorros
    rsX.Close

    strSQL = "select coalesce(sum(R.monto),0) as Monto" _
           & " from refundiciones R inner join reg_creditos C on R.id_solicitud = C.id_solicitud" _
           & " where R.id_solicitudr = " & Operacion.Operacion & " and C.garantia = 'A'"
    rsX.Open strSQL, glogon.Conection, adOpenStatic
    'Total de Saldos Sobre Ahorros menos las refundiciones sobre ahorros
    'Queda el pendiente en sobre ahorros
    vMontoRefunde = vMontoRefunde - rsX!Monto
    rsX.Close
    
    strSQL = "select coalesce(AHORRO,0) as Ahorro, coalesce(Capitaliza,0) as Capital" _
           & " from ahorro_consolidado where cedula = '" & txtCedula & "'"
    rsX.Open strSQL, glogon.Conection, adOpenStatic
    If (CCur(medMontoAprobado.Text) + vMontoRefunde) > CCur(Format(((rsX!Ahorro + rsX!capital) * (Porcentaje / 100)), "Standard")) Then
       vMensaje = vMensaje & vbCrLf & " - El Monto aprobado excede el " & Porcentaje _
                & "% de sus ahorros"
    End If
    rsX.Close
    
End If

'Verifica si la persona está o no bloqueda para formalizaciones
strSQL = "select bloqueo from socios where cedula = '" & Operacion.Cedula & "'"
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX!bloqueo = 1 Then
   vMensaje = vMensaje & vbCrLf & "- Esta persona se encuentra bloqueda, hasta mañana se le podran formalizar operaciones..."
End If
rsX.Close

'Ver si es Calculo de Excedente y Verificar el Monto aqui
strSQL = "select ase_codigo from excedentes_parametros"
rsX.Open strSQL, glogon.Conection, adOpenStatic
If UCase(Trim(rsX!ase_codigo)) = UCase(Trim(Operacion.Codigo)) Then
  If CCur(medMontoAprobado.Text) > CCur(Format(fxExcedenteDisponible(Operacion.Cedula), "Standard")) Then
     vMensaje = vMensaje & vbCrLf & "- Este es un prestamo sobre excedentes, y el monto aprobado sobrepasa la tabla autorizada..."
  End If
End If
rsX.Close



'Si es cheque verificar el número de Cheque a Registrar en Tesoreria
'y si el la transferencia del cheque a tesoreria es automatica
strSQL = "select PideCheque from catalogo where codigo = '" & Operacion.Codigo & "'"
rsX.Open strSQL, glogon.Conection, adOpenStatic
If Trim(UCase(cboTipoDocumento.Text)) = "CHEQUE" And rsX!pideCheque = "S" And Len(vMensaje) = 0 Then
   frmCR_SeguimientoDoc.Show vbModal
   If Not Operacion.Valida Then vMensaje = vMensaje & vbCrLf & " - Se registró inconsistencias en la especificación del número del documento a desembolsar."
End If 'CK
rsX.Close


If Len(vMensaje) > 0 Then fxVerificaFormalizacion = False


End Function

Private Function fxVerificaAnulacion() As Boolean
Dim rsX As New ADODB.Recordset, strSQL As String
Dim lngPrideduc As Long, vFecha As Date
Dim Porcentaje As Double, vMontoRefunde As Currency

vMensaje = ""
fxVerificaAnulacion = True

strSQL = "select COALESCE(count(*),0) as Existe" _
   & " from NIVEL_GRUPOS N INNER JOIN nivel_miembros A" _
   & " ON N.NV_COD_GRUPO = A.NV_COD_GRUPO INNER JOIN nivel_derechos B" _
   & " ON N.NV_COD_GRUPO = B.NV_COD_GRUPO Where A.nombre = '" & glogon.Usuario _
   & "' and B.codigo = '" & txtCodigo & "' AND N.nv_tipo = 'N'" _
   & " and (" & CCur(medMontoAprobado.Text) & " between nv_desde and nv_hasta)"

rsX.CursorLocation = adUseServer
rsX.Open strSQL, glogon.Conection, adOpenStatic
If rsX!existe = 0 Then
  vMensaje = vMensaje & vbCrLf & "- No existe nivel de anulación de este usuario para el código " & txtCodigo
End If
rsX.Close

If Len(vMensaje) > 0 Then fxVerificaAnulacion = False


End Function


Private Sub sbConsultaX(Vcedula As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

strSQL = "Select R.id_solicitud,R.codigo,R.cedula,R.fechasol,R.montosol,R.estadosol,R.estado,R.proceso" _
       & " FROM REG_CREDITOS R inner join CATALOGO C ON R.CODIGO = C.CODIGO" _
       & " where C.retencion = 'N' and C.poliza = 'N' and R.cedula = '" & Trim(txtConCedula) _
       & "' order by R.id_solicitud desc"

rs.CursorLocation = adUseServer
rs.Open strSQL, glogon.Conection, adOpenForwardOnly
lswBusca.ListItems.Clear
Do While Not rs.EOF
 Set itmX = lswBusca.ListItems.Add(, , CStr(rs!id_solicitud))
  itmX.SubItems(1) = rs!Codigo
  itmX.SubItems(2) = rs!Cedula
  itmX.SubItems(3) = Format(rs!fechasol, "yyyy/mm/dd")
  itmX.SubItems(4) = Format(rs!montosol, "###,###,###,##0.00")
  Select Case rs!estadosol
   Case "R"
    itmX.SubItems(5) = "Recibida"
   Case "P"
    itmX.SubItems(5) = "Pendiente"
   Case "A"
    itmX.SubItems(5) = "Aprobada"
   Case "D"
    itmX.SubItems(5) = "Denegada"
   Case "F"
    itmX.SubItems(5) = "Formalizada"
   Case "N"
    itmX.SubItems(5) = "Anulada"
  End Select
  
 Select Case rs!Estado
   Case "A"
    itmX.SubItems(6) = "Activa"
   Case "C"
    itmX.SubItems(6) = "Cancelada"
   Case Else
    itmX.SubItems(6) = "En Tramite"
 End Select
 
 Select Case rs!proceso
   Case "J"
    itmX.SubItems(7) = "Cobro Jud"
   Case "N"
    itmX.SubItems(7) = "Normal"
   Case "T"
    itmX.SubItems(7) = "Traspaso"
   Case Else
    itmX.SubItems(7) = "------"
 End Select
 
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault
Exit Sub

vError:
 Me.MousePointer = vbDefault
 MsgBox Err.Description, vbCritical

End Sub



Public Sub sbGXSegTraIniTlb()
 Call tlbPrincipal_ButtonClick(tlbPrincipal.Buttons.Item(1))
 txtCedula = frmCR_ConsultaCreditos.txtCedula
 txtCedula_LostFocus
 txtCodigo.SetFocus
End Sub

Private Sub Form_Activate()
vModulo = 3
End Sub

Private Sub Form_DblClick()
Set Conlsw.frmX = Me
Conlsw.ImprimeForm
End Sub

Private Sub Form_Load()
 
 vModulo = 3
 
 Call sbToolBarIconos(tlbPrincipal, False)
 
 Call Formularios(Me)
 
 
 Call LimpiaDatos
 Call sbCargaCombos
 With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With

Call RefrescaTags(Me)

End Sub

Private Sub LimpiaDatos()

With Operacion
 .Operacion = 0
 .Cedula = ""
 .Codigo = ""
 .EstadoSolicitud = "R"
 .Documento = ""
End With

 cboTipoDocumento.Text = "Transferencia"
 txtAmortizado = ""
 txtCedula = ""
 txtCodigo = ""
 txtCuentaAhorros = ""
 txtCuotaAprobada = ""
 txtCuotaSolicitada = ""
 txtDescripcion = ""
 txtInteresAprobado = ""
 txtInteresPagado = ""
 txtInteresSolicitado = ""
 txtNombre = ""
 txtObservaciones = ""
 txtPagare = ""
 txtPlazoAprobado = ""
 txtPlazoSolicitado = ""
 txtSaldo = ""
 medMontoAprobado = ""
 medMontoSolicitado = ""
 
 dtpFechaFormalizacion = fxFechaServidor
 dtpFechaRes = dtpFechaFormalizacion
 dtpFechaSolicitud = dtpFechaFormalizacion
 dtpDesembolso = dtpFechaFormalizacion
 
 txtProceso = ""
 txtEstadoActual = ""
 cboEstado.Clear
 cboGarantia.Clear
 cboDestino.Clear
 lblRecibe.Caption = ""
 lblResoluciona.Caption = ""
 lblFormaliza.Caption = ""
 lblTesoreria.Caption = ""
 
 chkEnviarATesoreria.Value = vbChecked
 chkPrimera.Value = vbChecked
 
 ssTabOperacion.Tab = 0
 ssTabOperacion.TabEnabled(0) = False
 ssTabOperacion.TabEnabled(1) = False
 ssTabOperacion.TabEnabled(2) = False
 ssTabOperacion.TabEnabled(3) = False

' Call sbCargaCombos

End Sub

Private Sub sbCargaCombos()
Dim strSQL As String, rs As New ADODB.Recordset


strSQL = "Select id_comite,descripcion from comites"
rs.Open strSQL, glogon.Conection, adOpenStatic
cboComite.Clear
 If rs.EOF And rs.BOF Then
   MsgBox "No existen Comités creados...(Debe Crearlos)", vbCritical
   Else
        Do While Not rs.EOF
         cboComite.AddItem rs!Descripcion & ""
         cboComite.ItemData(cboComite.NewIndex) = rs!id_comite
         rs.MoveNext
        Loop
        rs.MoveFirst
        cboComite.Text = rs!Descripcion
 End If
 rs.Close


cboBanco.Clear

strSQL = "select B.id_banco,B.descripcion" _
       & " from te_banco_asigna T inner join Bancos B on T.id_banco = B.id_banco" _
       & " where T.nombre = '" & glogon.Usuario & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
  MsgBox "No existen Bancos [Creados/Asignados], verifique en Tesoreria...", vbCritical

Else
 Do While Not rs.EOF
   cboBanco.AddItem IIf(IsNull(rs!Descripcion), "SIN DESCRIPCION", rs!Descripcion)
   cboBanco.ItemData(cboBanco.NewIndex) = rs!id_banco
   rs.MoveNext
 Loop
 rs.MoveFirst
 cboBanco.Text = IIf(IsNull(rs!Descripcion), "SIN DESCRIPCION", rs!Descripcion)
End If
rs.Close


cboTipoDocumento.Clear
cboTipoDocumento.AddItem fxTipoDocumento("RE")
cboTipoDocumento.AddItem fxTipoDocumento("CK")
cboTipoDocumento.AddItem fxTipoDocumento("TE")
cboTipoDocumento.AddItem fxTipoDocumento("ND")
'cboTipoDocumento.AddItem fxTipoDocumento("NC")

cboTipoDocumento.Text = fxTipoDocumento("TE")

'cboTipoDocumento.AddItem fxTipoDocumento("OT")

End Sub



Private Sub CargaRecursos(cbo As ComboBox, vCodigo As String, vGrupo As String)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vResultado As String

On Error Resume Next

cbo.Clear
vResultado = ""

strSQL = " select G.cod_grupo,(rtrim(G.cod_grupo) + ' - ' + rtrim(G.descripcion)) as Recurso" _
       & " from catalogo_grupos G inner join catalogo_asignaGrp A on G.cod_grupo = A.cod_grupo" _
       & " where A.codigo = '" & vCodigo & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
 If Trim(rs!cod_grupo) = Trim(vGrupo) Then vResultado = rs!recurso
 cbo.AddItem rs!recurso
 rs.MoveNext
Loop

If vResultado = "" Then
  If rs.RecordCount > 0 Then
     rs.MoveFirst
     cbo.Text = rs!recurso
  End If
Else
  cbo.Text = vResultado
End If
rs.Close

End Sub


Sub ActivaDesActiva(vEstadoSolicitud As String, vEstadoEC As String)
'Activa e Inactiva informacion, en los tabs

tlbFormalizacion.Enabled = True
optFormalizacion(0).Enabled = True
optFormalizacion(0).Value = True
cmdAplicarFormalizacion.Enabled = True


If vEstadoEC = "N" Then
    Select Case UCase(vEstadoSolicitud)
     Case "R"
       ssTabOperacion.TabEnabled(0) = True
       ssTabOperacion.TabEnabled(1) = True
       ssTabOperacion.TabEnabled(2) = False
       ssTabOperacion.TabEnabled(3) = False
       ssTabOperacion.Tab = 1
     Case "P"
       ssTabOperacion.TabEnabled(0) = True
       ssTabOperacion.TabEnabled(1) = False
       ssTabOperacion.TabEnabled(2) = False
       ssTabOperacion.TabEnabled(3) = False
       ssTabOperacion.Tab = 0
     Case "A"
       ssTabOperacion.TabEnabled(0) = False
       ssTabOperacion.TabEnabled(1) = True
       ssTabOperacion.TabEnabled(2) = True
       ssTabOperacion.TabEnabled(3) = False
       ssTabOperacion.Tab = 2
     Case "D"
       ssTabOperacion.TabEnabled(0) = True
       ssTabOperacion.TabEnabled(1) = True
       ssTabOperacion.TabEnabled(2) = False
       ssTabOperacion.TabEnabled(3) = False
       ssTabOperacion.Tab = 1
     Case "F"
       ssTabOperacion.TabEnabled(0) = False
       ssTabOperacion.TabEnabled(1) = False
       ssTabOperacion.TabEnabled(2) = True
       ssTabOperacion.TabEnabled(3) = False
       ssTabOperacion.Tab = 2
       tlbFormalizacion.Enabled = False
       optFormalizacion(1).Value = True
       optFormalizacion(0).Enabled = False
     Case "N"
       ssTabOperacion.TabEnabled(0) = False
       ssTabOperacion.TabEnabled(1) = False
       ssTabOperacion.TabEnabled(2) = True
       ssTabOperacion.TabEnabled(3) = False
       ssTabOperacion.Tab = 2
       cmdAplicarFormalizacion.Enabled = False
       tlbFormalizacion.Enabled = False
    End Select
Else
    Select Case UCase(vEstadoSolicitud)
      Case "F"
       ssTabOperacion.TabEnabled(0) = False
       ssTabOperacion.TabEnabled(1) = False
       ssTabOperacion.TabEnabled(2) = True
       ssTabOperacion.TabEnabled(3) = True
       ssTabOperacion.Tab = 2
       tlbFormalizacion.Enabled = False
       optFormalizacion(1).Value = True
       optFormalizacion(0).Enabled = False
      Case "N"
       ssTabOperacion.TabEnabled(0) = False
       ssTabOperacion.TabEnabled(1) = False
       ssTabOperacion.TabEnabled(2) = True
       ssTabOperacion.TabEnabled(3) = False
       cmdAplicarFormalizacion.Enabled = False
       tlbFormalizacion.Enabled = False
     End Select
'  ssTabOperacion.TabEnabled(0) = False
'  ssTabOperacion.TabEnabled(1) = False
'  ssTabOperacion.TabEnabled(2) = True
'  ssTabOperacion.TabEnabled(3) = True
  ssTabOperacion.Tab = 3
'  tlbFormalizacion.Enabled = False
'  optFormalizacion(0).Enabled = False
'  optFormalizacion(1).Value = True
End If
End Sub

Private Function fxBancoAsignado(vBanco As Integer, vUsuario As String) As Boolean
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "select coalesce(count(*),0) as Existe from te_banco_asigna where id_banco = " _
       & vBanco & " and nombre = '" & vUsuario & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs!existe = 0 Then
  fxBancoAsignado = False
Else
  fxBancoAsignado = True
End If
rs.Close

End Function

Private Function fxOperacionDestino(vDestino As String) As String
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select rtrim(cod_destino) + ' - ' + descripcion as ItemX from catalogo_destinos where cod_destino = '" & vDestino & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If rs.EOF And rs.BOF Then
  fxOperacionDestino = " -"
Else
  fxOperacionDestino = rs!itemX
End If
rs.Close

End Function

Private Sub sbCargaOperacion()
Dim rs As New ADODB.Recordset, strSQL As String
Dim vFecha As Date, iMes As Integer, lngAnio As Long
Dim i As Integer, vTemp As String

On Error Resume Next

strSQL = "select S.nombre,C.descripcion as CodDesc,X.descripcion as ComDesc," _
       & "R.cedula,R.codigo,R.id_solicitud,R.id_comite,R.Amortiza,R.saldo,R.estadosol,R.estado," _
       & "R.montoapr,R.ts,R.cuota,R.plazo,R.int,R.interesv,R.interesc,R.montosol," _
       & "R.observacion,R.pagare,R.fechaforp,R.fechasol,R.fechares,R.ind_deduce_planilla," _
       & "R.proceso,R.garantia,R.documento_referido,R.primer_cuota,R.emitir,R.prideduc," _
       & "R.userrec,R.userres,R.userfor,R.usertesoreria,R.cod_banco,R.CTA_BANCO" _
       & ",R.fecha_inicio_calculo,R.cod_grupo,R.cod_destino" _
       & " from reg_creditos R inner join Socios S on R.cedula = S.cedula" _
       & " inner join Catalogo C on R.codigo = C.codigo" _
       & " inner join Comites X on R.id_comite = X.id_comite" _
       & " where R.id_solicitud = " & txtOperacion

vFecha = fxFechaServidor

fraConsulta.Visible = False

rs.Open strSQL, glogon.Conection, adOpenStatic, adLockOptimistic

If Not rs.EOF And Not rs.BOF Then
  
 Call sbCargaCombos

 txtAmortizado = Format(IIf(IsNull(rs!Amortiza), 0, rs!Amortiza), "Standard")
 txtCedula = rs!Cedula
 txtNombre = rs!Nombre
 txtCodigo = rs!Codigo
 
 Operacion.Operacion = rs!id_solicitud
 Operacion.Cedula = rs!Cedula
 Operacion.Nombre = txtNombre
 Operacion.EstadoSolicitud = rs!estadosol
 Operacion.Codigo = rs!Codigo
 Operacion.Estado = IIf(IsNull(rs!Estado), "N", rs!Estado)
 Operacion.MontoAprobado = IIf(IsNull(rs!montoApr), 0, rs!montoApr)
 Operacion.TS = fxTsToHex(rs!TS)  'TimeStamp
  
' MsgBox Operacion.TS & vbCrLf & rs!TS
 txtDescripcion = rs!CodDesc
 txtCuentaAhorros = IIf(IsNull(rs!cta_banco), "", rs!cta_banco)
 
 txtCuotaAprobada = Format(IIf(IsNull(rs!cuota), 0, rs!cuota), "Standard")
 txtCuotaSolicitada = Format(IIf(IsNull(rs!cuota), 0, rs!cuota), "Standard")
 txtPlazoAprobado = IIf(IsNull(rs!Plazo), 0, rs!Plazo)
 txtPlazoSolicitado = IIf(IsNull(rs!Plazo), 0, rs!Plazo)
 txtInteresAprobado = IIf(IsNull(rs!Int), 0, rs!Int)
 txtInteresSolicitado = IIf(IsNull(rs!Int), 0, rs!Int)
 txtInteresPagado = Format(IIf(IsNull(rs!interesc), 0, rs!interesc), "Standard")
 If rs!estadosol = "P" Or rs!estadosol = "R" Then
   medMontoAprobado = Format(IIf(IsNull(rs!montosol), 0, rs!montosol), "Standard")
 Else
   medMontoAprobado = Format(IIf(IsNull(rs!montoApr), 0, rs!montoApr), "Standard")
 End If
 medMontoSolicitado = Format(IIf(IsNull(rs!montosol), 0, rs!montosol), "Standard")
 
 txtObservaciones = IIf(IsNull(rs!observacion), "", rs!observacion)
 txtPagare = IIf(IsNull(rs!pagare), 0, rs!pagare)
 txtSaldo = Format(IIf(IsNull(rs!Saldo), 0, rs!Saldo), "Standard")
 dtpFechaFormalizacion = IIf(IsNull(rs!fechaforp), vFecha, rs!fechaforp)
 dtpFechaRes = IIf(IsNull(rs!fechares), vFecha, rs!fechares)
 dtpDesembolso = IIf(IsNull(rs!fecha_inicio_calculo), vFecha, rs!fecha_inicio_calculo)
 dtpFechaSolicitud = IIf(IsNull(rs!fechasol), vFecha, rs!fechasol)
 txtEstadoActual = fxDescribeEstado(IIf(IsNull(rs!Estado), "N", rs!Estado))
 cboComite.Text = rs!Comdesc
  
 'Carga Destino
 Call sbSTCargaCboDestinos(cboDestino, Operacion.Codigo)
 vTemp = fxOperacionDestino(rs!cod_destino & "")
 cboDestino.AddItem vTemp
 cboDestino.Text = vTemp
  
 'Si no tiene el banco asignado hay que crearlo pero no puede guardarlo
 'bajo este mismo banco hasta que lo tenga asignado o lo cambie.
 If Not fxBancoAsignado(rs!Cod_Banco, glogon.Usuario) Then
     cboBanco.AddItem fxDescribeBanco(IIf(IsNull(rs!Cod_Banco), 0, rs!Cod_Banco))
     cboBanco.ItemData(cboBanco.NewIndex) = rs!Cod_Banco
 End If
 cboBanco.Text = fxDescribeBanco(IIf(IsNull(rs!Cod_Banco), 0, rs!Cod_Banco))
 
 Select Case IIf(IsNull(rs!proceso), "N", rs!proceso)
  Case "N"
    txtProceso = "NORMAL"
  Case "T"
    txtProceso = "TRA.DEUDAS"
  Case "J"
    txtProceso = "CBR.JUDICIAL"
  Case Else
    txtProceso = ""
 End Select
 
 chkDeducirPlanilla.Value = IIf((rs!ind_deduce_planilla = "S"), 1, 0)
 chkPrimera.Value = IIf((rs!PRIMER_CUOTA = "S"), 1, 0)
 '**
 txtDocumento = IIf(IsNull(rs!documento_referido), "", rs!documento_referido)
 cboTipoDocumento.Text = fxTipoDocumento(IIf(IsNull(rs!emitir), "OT", rs!emitir))
 
 
 If IsNull(rs!prideduc) Then
  
  cboMes.Text = fxConvierteMES(Val(Mid(fxPrimerDeduccion, 5, 2)))
  txtAno.Text = Mid(fxPrimerDeduccion, 1, 4)
  
  If chkPrimera.Value = vbChecked Then
    cboMes.Text = fxConvierteMES(Val(Mid(fxPrimerDeduccionCuota, 5, 2)))
    txtAno.Text = Mid(fxPrimerDeduccionCuota, 1, 4)
  End If 'Primera Cuota Marcada
   
 Else
 
  cboMes.Text = fxConvierteMES(Val(Mid(rs!prideduc, 5, 2)))
  txtAno.Text = Mid(rs!prideduc, 1, 4)
 
 End If
 
 Call sbSTCargaCboEstado(cboEstado, rs!estadosol)
 Call sbSTCargaCboGarantia(cboGarantia, rs!Codigo)
 Call CargaRecursos(cboRecursos, rs!Codigo, rs!cod_grupo & "")
 
 cboGarantia.Text = fxGarantia(rs!garantia)
 
 Call ActivaDesActiva(rs!estadosol, IIf(IsNull(rs!Estado), "N", rs!Estado))

 lblRecibe.Caption = IIf(IsNull(rs!userrec), "", rs!userrec)
 lblResoluciona.Caption = IIf(IsNull(rs!userres), "", rs!userres)
 lblFormaliza.Caption = IIf(IsNull(rs!userfor), "", rs!userfor)
 lblTesoreria.Caption = IIf(IsNull(rs!usertesoreria), "", rs!usertesoreria)

 With tlbPrincipal.Buttons
   .Item(1).Enabled = True
   .Item(2).Enabled = True
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With
 Me.fraOperacion.Enabled = False
Else
 MsgBox "No existe esta Solicitud", vbCritical
End If
rs.Close

Call RefrescaTags(Me)

End Sub

Private Sub sbBusqueda(Index As Integer)
'Set GLOBALES.gfrmFormulario = Me
gBusquedas.Resultado = ""
gBusquedas.Convertir = "N"

Select Case Index
  Case 0 'txtOperacion
    gBusquedas.Convertir = "S"
    gBusquedas.Consulta = "select id_solicitud as Operacion,codigo,cedula,montoapr,saldo from reg_creditos"
    gBusquedas.Orden = "id_solicitud"
    gBusquedas.Columna = "id_solicitud"
    frmBusquedas.Show vbModal
    txtOperacion = gBusquedas.Resultado
    If Len(Trim(txtOperacion)) > 0 Then
    '  Call ConsultaOperacion
    End If
  
  Case 1 'txtCedula
        gBusquedas.Consulta = "select Cedula,Nombre from socios"
        gBusquedas.Orden = "cedula"
        gBusquedas.Columna = "cedula"
        frmBusquedas.Show vbModal
       ' Call CambiaDatos
        txtCedula = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          txtNombre = fxNombre(Trim(txtCedula))
        End If
  
  
  Case 2 'txtCodigo
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "codigo"
        gBusquedas.Columna = "codigo"
        gBusquedas.Filtro = " and Activo = 1 and Retencion = 'N'"
        frmBusquedas.Show vbModal
       ' Call CambiaDatos
        txtCodigo = gBusquedas.Resultado
        If Len(Trim(txtCodigo)) > 0 Then
          txtDescripcion = fxDescribeCodigo(Trim(txtCodigo))
        End If
  
  Case 3 'txtNombre
        gBusquedas.Consulta = "select Cedula,Nombre from socios"
        gBusquedas.Orden = "nombre"
        gBusquedas.Columna = "nombre"
        frmBusquedas.Show vbModal
       ' Call CambiaDatos
        txtCedula = gBusquedas.Resultado
        If Len(Trim(txtCedula)) > 0 Then
          txtNombre = fxNombre(Trim(txtCedula))
        End If
  
  
  Case 4 'txtDescripcion
        gBusquedas.Consulta = "select codigo,descripcion from catalogo"
        gBusquedas.Orden = "descripcion"
        gBusquedas.Columna = "descripcion"
        frmBusquedas.Show vbModal
       ' Call CambiaDatos
        txtCodigo = gBusquedas.Resultado
        If Len(Trim(txtCodigo)) > 0 Then
          txtDescripcion = fxDescribeCodigo(Trim(txtCodigo))
        End If
  
End Select

End Sub







Private Sub imgConsulta_Click()
fraConsulta.Visible = True
fraConsulta.Top = 1080
lswBusca.ListItems.Clear
txtConCedula = ""
txtConNombre = ""
End Sub

Private Sub imgConsultaCerrar_Click()
fraConsulta.Visible = False
End Sub


Private Sub imgConsultaUsers_Click()
fraUser.Visible = False
End Sub

Private Sub imgGuardaFecDesembolso_Click()
Dim strSQL As String

strSQL = "update reg_creditos set fecha_inicio_calculo = '" & Format(dtpDesembolso.Value, "yyyy/mm/dd") & "'" _
       & " where id_solicitud = " & Operacion.Operacion & " and estadosol not in('F','N')"
glogon.Conection.Execute strSQL

Call Bitacora("Modifica", "Fecha de Desembolso Operacion " & Operacion.Operacion)

MsgBox "Fecha de Desembolso Actualizada satisfactoriamente...", vbInformation

End Sub

Private Sub imgRecibir_Click()
Dim strSQL As String

strSQL = "update reg_creditos set estadosol = 'R' where id_solicitud = " _
       & Operacion.Operacion & " and estadosol not in('F','N')"
glogon.Conection.Execute strSQL

Call Bitacora("Aplica", "Pone Como Recibida la Op:" & Operacion.Operacion)

Call sbCargaOperacion

End Sub

Private Sub imgRequisitos_Click()
Call cmdRequisitos_Click
End Sub

Private Sub lswBusca_Click()
If lswBusca.ListItems.Count > 0 Then
 txtOperacion = lswBusca.SelectedItem
 Call sbCargaOperacion
End If
End Sub

Private Sub medMontoSolicitado_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn And medMontoSolicitado.Text <> "" Then
   txtPlazoSolicitado.Text = fxCatalogoRango(txtCodigo, medMontoSolicitado.Text, "P")
   txtInteresSolicitado.Text = fxCatalogoRango(txtCodigo, medMontoSolicitado.Text, "I", fxCodigoCbo(cboDestino))
   txtPlazoSolicitado.SetFocus
 End If
End Sub



Private Sub Edicion(intActiva As Integer)
'Activa e inactiva partes a editar

If intActiva = 1 Then
  fraOperacion.Enabled = True
  Select Case Operacion.EstadoSolicitud
   Case "R", "P"
     Me.txtCedula.Enabled = True
     Me.txtCodigo.Enabled = True
'     Me.imgBusqueda_Rapida(1).Enabled = True
'     Me.imgBusqueda_Rapida(2).Enabled = True
     Me.medMontoSolicitado.Enabled = True
     Me.txtPlazoSolicitado.Enabled = True
     Me.txtInteresSolicitado.Enabled = True
     Me.cboBanco.Enabled = True
     Me.cboComite.Enabled = True
     Me.cboGarantia.Enabled = True
     Me.cboEstado.Enabled = True
     Me.txtObservaciones.Enabled = True
     Me.txtCuentaAhorros.Enabled = True
'     Me.imgBusqueda_Rapida(3).Enabled = False
'     Me.imgBusqueda_Rapida(4).Enabled = False
     Me.dtpFechaSolicitud.Enabled = True
     Me.cboTipoDocumento.Enabled = True

   Case "A"
     Me.txtCedula.Enabled = True
     Me.txtCodigo.Enabled = True
'     Me.imgBusqueda_Rapida(2).Enabled = True
     Me.cboBanco.Enabled = True
     Me.cboComite.Enabled = True
     Me.cboGarantia.Enabled = True
     Me.txtObservaciones.Enabled = True
     Me.txtCuentaAhorros.Enabled = True
     Me.cboTipoDocumento.Enabled = True
     
'     Me.imgBusqueda_Rapida(1).Enabled = False
'     Me.imgBusqueda_Rapida(3).Enabled = False
'     Me.imgBusqueda_Rapida(4).Enabled = False
     Me.dtpFechaSolicitud.Enabled = False
     Me.medMontoSolicitado.Enabled = False
     Me.txtPlazoSolicitado.Enabled = False
     Me.txtInteresSolicitado.Enabled = False
     Me.cboEstado.Enabled = False
   Case "D", "N"
     Me.txtObservaciones.Enabled = True
     Me.txtCedula.Enabled = False
     Me.txtCodigo.Enabled = False
     Me.cboBanco.Enabled = False
     Me.cboComite.Enabled = False
     Me.cboGarantia.Enabled = False
     Me.txtCuentaAhorros.Enabled = False
'     Me.imgBusqueda_Rapida(1).Enabled = False
'     Me.imgBusqueda_Rapida(2).Enabled = False
'     Me.imgBusqueda_Rapida(3).Enabled = False
'     Me.imgBusqueda_Rapida(4).Enabled = False
     Me.dtpFechaSolicitud.Enabled = False
     Me.medMontoSolicitado.Enabled = False
     Me.txtPlazoSolicitado.Enabled = False
     Me.txtInteresSolicitado.Enabled = False
     Me.cboEstado.Enabled = False
     Me.cboTipoDocumento.Enabled = False
   Case "F"
     Me.txtObservaciones.Enabled = True
     Me.txtCedula.Enabled = False
     Me.txtCodigo.Enabled = False
     Me.cboBanco.Enabled = True
     Me.cboComite.Enabled = False
     Me.cboGarantia.Enabled = False
     Me.txtCuentaAhorros.Enabled = True
'     Me.imgBusqueda_Rapida(1).Enabled = False
'     Me.imgBusqueda_Rapida(2).Enabled = False
'     Me.imgBusqueda_Rapida(3).Enabled = False
'     Me.imgBusqueda_Rapida(4).Enabled = False
     Me.dtpFechaSolicitud.Enabled = False
     Me.medMontoSolicitado.Enabled = False
     Me.txtPlazoSolicitado.Enabled = False
     Me.txtInteresSolicitado.Enabled = False
     Me.cboEstado.Enabled = False
     Me.cboTipoDocumento.Enabled = True
  End Select
Else 'apaga
  fraOperacion.Enabled = False
     Me.txtCedula.Enabled = True
     Me.txtCodigo.Enabled = True
'     Me.imgBusqueda_Rapida(1).Enabled = True
'     Me.imgBusqueda_Rapida(2).Enabled = True
'     Me.imgBusqueda_Rapida(3).Enabled = True
'     Me.imgBusqueda_Rapida(4).Enabled = True
     Me.medMontoSolicitado.Enabled = True
     Me.txtPlazoSolicitado.Enabled = True
     Me.txtInteresSolicitado.Enabled = True
     Me.cboBanco.Enabled = True
     Me.cboComite.Enabled = True
     Me.cboGarantia.Enabled = True
     Me.cboEstado.Enabled = True
     Me.txtObservaciones.Enabled = True
     Me.txtCuentaAhorros.Enabled = True
     Me.dtpFechaSolicitud.Enabled = True
     Me.cboTipoDocumento.Enabled = True
  Select Case Operacion.EstadoSolicitud
   Case "A"
   Case "D"
   Case "N"
   Case "F"
  End Select
End If 'inactiva
End Sub

Private Sub sbCargosAdicionales(vOperacion As Long, vCodigo As String, vMonto As Currency)
Dim strSQL As String, rs As New ADODB.Recordset

strSQL = "delete operacion_cargos where id_solicitud = " & vOperacion
glogon.Conection.Execute strSQL

strSQL = "select * from cargos_asignacion where codigo = '" & vCodigo & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
Do While Not rs.EOF
  strSQL = "insert operacion_cargos(id_solicitud,codigo,cod_cargo,tipo,valor,monto)" _
         & " values(" & vOperacion & ",'" & vCodigo & "','" & rs!cod_cargo & "','" _
         & rs!Tipo & "'," & rs!valor
  If rs!Tipo = "P" Then
    strSQL = strSQL & (vMonto * (rs!valor / 100)) & ")"
  Else
    strSQL = strSQL & (vMonto + rs!valor) & ")"
  End If
  rs.MoveNext
Loop
rs.Close

End Sub


Private Sub sbGuardarSolicitud()
Dim strSQL As String, rs As New ADODB.Recordset

If vEdita Then
  Select Case Operacion.EstadoSolicitud
    Case "R", "P"
      
      Call sbCargosAdicionales(Operacion.Operacion, txtCodigo, CCur(medMontoSolicitado))
       
      strSQL = "update reg_creditos set cedula =" _
         & "'" & Trim(txtCedula) & "',codigo = '" & Trim(txtCodigo) & "',montosol=" & CCur(medMontoSolicitado.Text) _
         & ",fechasol='" & Format(Me.dtpFechaSolicitud.Value, "yyyy/mm/dd") & "',estadosol='" _
         & fxEstadoOperacion(cboEstado.Text) & "',id_comite=" & cboComite.ItemData(cboComite.ListIndex) & ",int=" _
         & Trim(txtInteresSolicitado) & ",interesv=" & Trim(txtInteresSolicitado) & ",plazo=" _
         & Trim(txtPlazoSolicitado) & ",cuota=" & CCur(txtCuotaSolicitada) & ",garantia='" _
         & fxGarantia(cboGarantia.Text) & "',observacion=" _
         & IIf((Len(Trim(txtObservaciones)) = 0), "'NADA'", "'" & Mid(Trim(txtObservaciones), 1, 254) & "'") _
         & ",acta =null,estado =null,userrec='" & glogon.Usuario & "',cod_banco=" & cboBanco.ItemData(cboBanco.ListIndex) & ",cta_banco=" _
         & IIf((Len(Trim(txtCuentaAhorros)) = 0), "0", Mid(Trim(txtCuentaAhorros), 1, 20)) _
         & ",ind_deposito=" & IIf(UCase(cboTipoDocumento.Text) = "TRANSFERENCIA", 1, 0) _
         & ",emitir='" & fxTipoDocumento(cboTipoDocumento.Text) & "'" _
         & ",primer_cuota ='" & IIf((chkPrimera.Value = vbChecked), "S", "N") & "',premio=0,tdocumento='ND'" _
         & ",cod_destino = '" & fxCodigoCbo(cboDestino) & "'"
    Case "A"
      
      Call sbCargosAdicionales(Operacion.Operacion, txtCodigo, CCur(medMontoAprobado))
      
      strSQL = "update reg_creditos set codigo = '" & Trim(txtCodigo) & "'" _
         & ",id_comite=" & cboComite.ItemData(cboComite.ListIndex) & "," _
         & "garantia='" & fxGarantia(cboGarantia.Text) & "',observacion=" _
         & IIf((Len(Trim(txtObservaciones)) = 0), "'NADA'", "'" & Mid(Trim(txtObservaciones), 1, 254) & "'") _
         & ",estado =null,cod_banco=" & cboBanco.ItemData(cboBanco.ListIndex) & ",cta_banco=" _
         & IIf((Len(Trim(txtCuentaAhorros)) = 0), "0", Mid(Trim(txtCuentaAhorros), 1, 20)) _
         & ",ind_deposito=" & IIf(UCase(cboTipoDocumento.Text) = "TRANSFERENCIA", 1, 0) _
         & ",emitir='" & fxTipoDocumento(cboTipoDocumento.Text) & "'" _
         & ",primer_cuota ='" & IIf((chkPrimera.Value = vbChecked), "S", "N") & "'"
         
    Case "F"
      strSQL = "update reg_creditos set observacion=" _
         & IIf((Len(Trim(txtObservaciones)) = 0), "'NADA'", "'" & Mid(Trim(txtObservaciones), 1, 254) & "'") _
         & ",cod_banco=" & cboBanco.ItemData(cboBanco.ListIndex) & ",cta_banco=" _
         & IIf((Len(Trim(txtCuentaAhorros)) = 0), "0", Mid(Trim(txtCuentaAhorros), 1, 20)) _
         & ",ind_deposito=" & IIf(UCase(cboTipoDocumento.Text) = "TRANSFERENCIA", 1, 0) _
         & ",emitir='" & fxTipoDocumento(cboTipoDocumento.Text) & "'"
    
    Case "N", "D"
      strSQL = "update reg_creditos set observacion=" _
         & IIf((Len(Trim(txtObservaciones)) = 0), "'NADA'", "'" & Mid(Trim(txtObservaciones), 1, 254) & "'")
  
  End Select
  strSQL = strSQL & " Where id_solicitud = " & txtOperacion
'         & " and TS = " & Operacion.TS
  glogon.Conection.Execute strSQL
    
  Call Bitacora("Registra", "Actualiza la Solicitud : " & Operacion.Operacion)
  MsgBox "Solicitud Actualizada Satisfactoriamente...", vbInformation

Else 'Inserta
  
  chkPrimera.Value = IIf(fxPrimerCuota(txtCodigo), vbChecked, vbUnchecked)
  
  strSQL = "insert into reg_creditos(cedula,codigo,montosol,fechasol,estadosol,id_comite," _
         & "int,interesv,plazo,cuota,garantia,observacion,estado,userrec,cod_banco,cta_banco," _
         & "ind_deposito,primer_cuota,premio,emitir,tdocumento,cod_destino) values(" _
         & "'" & Trim(txtCedula) & "','" & Trim(txtCodigo) & "'," & CCur(medMontoSolicitado.Text) _
         & ",'" & Format(dtpFechaSolicitud.Value, "yyyy/mm/dd") & "','" _
         & fxEstadoOperacion(cboEstado.Text) & "'," & cboComite.ItemData(cboComite.ListIndex) & "," _
         & Trim(txtInteresSolicitado) & "," & Trim(txtInteresSolicitado) & "," _
         & Trim(txtPlazoSolicitado) & "," & CCur(txtCuotaSolicitada) & ",'" _
         & fxGarantia(cboGarantia.Text) & "'," _
         & IIf((Len(Trim(txtObservaciones)) = 0), "'NADA'", "'" & Mid(Trim(txtObservaciones), 1, 254) & "'") _
         & ",null,'" & glogon.Usuario & "'," & cboBanco.ItemData(cboBanco.ListIndex) & "," _
         & IIf((Len(Trim(txtCuentaAhorros)) = 0), "0", Mid(Trim(txtCuentaAhorros), 1, 20)) _
         & "," & IIf(UCase(cboTipoDocumento.Text) = "TRANSFERENCIA", 1, 0) & ",'" _
         & IIf((chkPrimera.Value = vbChecked), "S", "N") & "',0,'" _
         & fxTipoDocumento(cboTipoDocumento.Text) & "','ND','" & fxCodigoCbo(cboDestino) & "')"
   'Verificar si existe la cuenta de ahorros y si no crearla
 glogon.Conection.Execute strSQL
 txtOperacion.Text = fxUltimaOperacion(txtCedula)
 
 Call sbCargosAdicionales(txtOperacion, txtCodigo, CCur(medMontoSolicitado))
  
 Call Bitacora("Registra", "Recepción de la Operacion : " & txtOperacion.Text)
 MsgBox "Solicitud Grabada Satisfactoriamente...", vbInformation
End If


End Sub

Private Sub ActualizaCodigoOperacion()
Dim strSQL As String

strSQL = "update fiadores set codigo = '" & txtCodigo & "' where id_solicitud =" & Operacion.Operacion
glogon.Conection.Execute strSQL

strSQL = "update refundiciones set codigor = '" & txtCodigo & "' where id_solicitudr =" & Operacion.Operacion
glogon.Conection.Execute strSQL

strSQL = "update desembolsos set codigo = '" & txtCodigo & "' where id_solicitud =" & Operacion.Operacion
glogon.Conection.Execute strSQL

strSQL = "update pra_principal set codigo = '" & txtCodigo & "' where id_solicitud =" & Operacion.Operacion
glogon.Conection.Execute strSQL



End Sub



Private Sub tlbAux_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim X As clsPreaAnalisis

Select Case Button.Key
  Case "calculo"
    frmCR_CalculoOperacion.Show
    frmCR_CalculoOperacion.txtCedula = Operacion.Cedula
  
  Case "preanalisis"
    Set X = New clsPreaAnalisis
    Set X.vCon = glogon.Conection
    X.xOperacion = Operacion.Operacion
    
    Call X.sbShow(glogon.Usuario, App.Path, 2)
    
    Set X = Nothing
  Case "usersx"
    fraUser.Visible = True
End Select

End Sub

Private Sub tlbFormalizacion_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim strSQL As String, rs As New ADODB.Recordset

Select Case Button.Key
 
 Case "refundiciones"
  strSQL = "select refunde from catalogo where codigo = '" & Operacion.Codigo & "'"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  If rs!refunde = "S" Then
    frmCR_SeguimientoRefundiciones.Show vbModal
  Else
    MsgBox "Esta Línea No Permite que se realicen refundiciones con ella...", vbCritical
  End If
  rs.Close
 
 Case "desembolsos"
  frmCR_SeguimientoDesembolsos.Show vbModal
 
 Case "retenciones"
  strSQL = "select refunde from catalogo where codigo = '" & Operacion.Codigo & "'"
  rs.Open strSQL, glogon.Conection, adOpenStatic
  If rs!refunde = "S" Then
    frmCR_SeguimientoRetenciones.Show vbModal
  Else
    MsgBox "Esta Línea No Permite que se realicen refundiciones con ella...", vbCritical
  End If
  rs.Close
  
 Case "firmas"
   frmCR_SeguimientoFirmas.Show vbModal
 
 Case "requisitos"
   Operacion.Ventana = "R"
   frmCR_SeguimientoReqCar.Show vbModal
 
 Case "cargos"
   Operacion.Ventana = "C"
   frmCR_SeguimientoReqCar.Show vbModal
End Select

End Sub

Private Sub tlbPrincipal_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

Select Case Button.Key
 Case "nuevo"
  txtOperacion.Text = ""
  txtOperacion.Enabled = False
'  imgBusqueda_Rapida(0).Enabled = False
  Call LimpiaDatos
  tlbPrincipal.Buttons(1).Enabled = False
  tlbPrincipal.Buttons(2).Enabled = False
  tlbPrincipal.Buttons(3).Enabled = True
  tlbPrincipal.Buttons(4).Enabled = True
  fraOperacion.Enabled = True
  vEdita = False
  txtCedula.SetFocus
  
  Call sbCargaCombos
  
  
 Case "editar"
  If Operacion.Operacion > 0 Then 'And Operacion.Estado = "A" Then
      vEdita = True
      Call Edicion(1)
    
      'Si el Estado Esta en Recepcion o Resolucion puede Cambiar Todos Los Datos
      'Si Esta en Formalización Solo puede Cambiar la Salida
      tlbPrincipal.Buttons(1).Enabled = False
      tlbPrincipal.Buttons(2).Enabled = False
      tlbPrincipal.Buttons(3).Enabled = True
      tlbPrincipal.Buttons(4).Enabled = True
      txtOperacion.Enabled = False
'      imgBusqueda_Rapida(0).Enabled = False
      fraOperacion.Enabled = True
      txtCedula.SetFocus

  End If
 
 Case "guardar"
  
  If fxVerificaRecepcion Then
    'Verificar si se cambio el codigo
    If Trim(txtCodigo) <> Operacion.Codigo Then Call ActualizaCodigoOperacion
    Call sbGuardarSolicitud
    Call Edicion(0)
    Call sbCargaOperacion
    txtOperacion.Enabled = True
    
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    
    fraOperacion.Enabled = False
    
    If vEdita = False Then
        ssTabOperacion.Tab = 0
        cmdDatosPersonales_Click
    End If
    If vEdita = False And cboGarantia.Text = "Fiadores" Then
        ssTabOperacion.Tab = 0
        cmdFiadores_Click
    End If
    If vEdita = False Then
        cmdRequisitos_Click
    End If
  
  
  Else
    MsgBox vMensaje, vbCritical
  End If
 
 Case "deshacer"
    txtOperacion.Enabled = True
    tlbPrincipal.Buttons(1).Enabled = True
    tlbPrincipal.Buttons(2).Enabled = True
    tlbPrincipal.Buttons(3).Enabled = False
    tlbPrincipal.Buttons(4).Enabled = False
    fraOperacion.Enabled = False
    If txtOperacion <> "" Then Call sbCargaOperacion
    txtOperacion.SetFocus
 
 Case "ayuda"
        frmContenedor.CD.HelpContext = Me.HelpContextID
        frmContenedor.CD.ShowHelp
 
 Case "cerrar"
  Unload Me

End Select


End Sub

Public Sub ReporteBoleta()
Dim strRuta As String, rs As New ADODB.Recordset

strRuta = App.Path + "\credito\reportes\crBoletaFormalizacion.rpt"
Me.MousePointer = vbHourglass

With frmContenedor.Crt
 .Reset
 .WindowShowRefreshBtn = True
 .WindowShowPrintSetupBtn = True
 .WindowState = crptMaximized
 .WindowShowSearchBtn = True
 .WindowTitle = "Boleta de Formalización"
 .ReportFileName = strRuta
 .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD}=" & Operacion.Operacion
 .Formulas(0) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy HH:MM:ss") & "'"
 .Formulas(1) = "Usuario='" & UCase(glogon.Usuario) & "'"
 .Formulas(2) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
 .PrintReport
End With

Me.MousePointer = vbDefault

End Sub

Private Sub tlbPrincipal_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Me.MousePointer = vbHourglass

Select Case ButtonMenu.Key
 Case "RepActas"
   frmCR_SolCreacionAgenda.Show vbModal
 Case "RepPreAnalisis"
   frmCR_SolicitudesPreAnalisis.Show vbModal
 Case "RepGarantia"
   frmCR_GeneraGarantia.Show vbModal
 Case "repBoleta"
   If Operacion.EstadoSolicitud = "F" Or Operacion.EstadoSolicitud = "N" Then
     Call ReporteBoleta
   Else
     MsgBox "La Operación # " & Operacion.Operacion & " No se encuentra formalizada", vbInformation
   End If
End Select

Me.MousePointer = vbDefault

End Sub

Private Sub txtCedula_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(1)
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtCodigo.SetFocus
End Sub

Private Sub txtCedula_LostFocus()
 txtNombre = fxNombre(txtCedula)
End Sub

Private Sub DescribeCodigoComite()
Dim rs As New ADODB.Recordset, strSQL As String

strSQL = "select coalesce(id_comite,0) as id_comite from catalogo where codigo='" & txtCodigo & "'"
rs.Open strSQL, glogon.Conection, adOpenStatic
If Not rs.EOF And Not rs.BOF Then
  cboComite.Text = fxDescribeComite(rs!id_comite)
End If
rs.Close

End Sub

Private Sub txtCodigo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(2)
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
On Error Resume Next

If KeyAscii = vbKeyReturn Then
  txtCodigo = UCase(txtCodigo)
  Call sbSTCargaCboGarantia(cboGarantia, txtCodigo)
  Call sbSTCargaCboEstado(cboEstado, "R")
  Call sbSTCargaCboDestinos(cboDestino, txtCodigo)
  
  If UCase(cboGarantia.Text) = "FONDOS / PLANES" Then
    medMontoSolicitado = fxDisponibleFondos(txtCedula)
  End If
  
  If fxCreditoExcedente(txtCodigo) Then
    chkPrimera.Value = vbUnchecked
    medMontoSolicitado.Text = fxExcedenteDisponible(txtCedula)
  End If
  
  cboDestino.SetFocus

End If
End Sub

Private Sub txtCodigo_LostFocus()
 txtDescripcion = fxDescribeCodigo(txtCodigo)
 Call DescribeCodigoComite
End Sub

Private Sub txtConCedula_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  txtConNombre = fxNombre(txtConCedula)
  Call sbConsultaX(txtConCedula)
End If

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Cedula"
  gBusquedas.Orden = "Cedula"
  gBusquedas.Consulta = "select Cedula,Nombre from socios"
  frmBusquedas.Show vbModal
  txtConCedula = gBusquedas.Resultado
  txtConNombre = gBusquedas.Resultado2
  Call sbConsultaX(txtConCedula)
End If

End Sub

Private Sub txtConNombre_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then Call sbConsultaX(txtConCedula)

If KeyCode = vbKeyF4 Then
  gBusquedas.Convertir = "N"
  gBusquedas.Columna = "Nombre"
  gBusquedas.Orden = "Nombre"
  gBusquedas.Consulta = "select Cedula,Nombre from socios"
  frmBusquedas.Show vbModal
  txtConCedula = gBusquedas.Resultado
  txtConNombre = gBusquedas.Resultado2
  Call sbConsultaX(txtConCedula)
End If

End Sub

Private Sub txtCuentaAhorros_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then txtObservaciones.SetFocus
End Sub

Private Sub txtDescripcion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(4)
End Sub

Private Sub txtInteresAprobado_LostFocus()
On Error GoTo vError
If CCur(IIf((txtInteresAprobado = ""), 0, txtInteresAprobado)) >= 0 And CCur(IIf((txtPlazoAprobado = ""), 0, txtPlazoAprobado)) > 0 _
    And CCur(IIf((medMontoAprobado = ""), 0, medMontoAprobado)) > 0 Then
  txtCuotaAprobada = fxCalcula_Cuota(CCur(medMontoAprobado), CCur(txtPlazoAprobado), CCur(txtInteresAprobado))
End If
vError:
End Sub

Private Sub txtInteresSolicitado_Change()
On Error GoTo vError
If CCur(IIf((txtInteresSolicitado = ""), 0, txtInteresSolicitado)) >= 0 And CCur(IIf((txtPlazoSolicitado = ""), 0, txtPlazoSolicitado)) > 0 _
    And CCur(IIf((medMontoSolicitado = ""), 0, medMontoSolicitado)) > 0 Then
  txtCuotaSolicitada = fxCalcula_Cuota(CCur(medMontoSolicitado), CCur(txtPlazoSolicitado), CCur(txtInteresSolicitado))
End If
vError:

End Sub

Private Sub medMontoSolicitado_Change()
On Error GoTo vError
If CCur(IIf((txtInteresSolicitado = ""), 0, txtInteresSolicitado)) > 0 And CCur(IIf((txtPlazoSolicitado = ""), 0, txtPlazoSolicitado)) > 0 _
    And CCur(IIf((medMontoSolicitado = ""), 0, medMontoSolicitado)) > 0 Then
 txtCuotaSolicitada = fxCalcula_Cuota(CCur(medMontoSolicitado), CCur(txtPlazoSolicitado), CCur(txtInteresSolicitado))
End If

vError:
End Sub

Private Sub txtInteresSolicitado_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then
  On Error Resume Next
    If CCur(IIf((txtInteresSolicitado = ""), 0, txtInteresSolicitado)) >= 0 And CCur(IIf((txtPlazoSolicitado = ""), 0, txtPlazoSolicitado)) > 0 _
        And CCur(IIf((medMontoSolicitado = ""), 0, medMontoSolicitado)) > 0 Then
      txtCuotaSolicitada = fxCalcula_Cuota(CCur(medMontoSolicitado), CCur(txtPlazoSolicitado), CCur(txtInteresSolicitado))
    End If
   cboGarantia.SetFocus
End If
End Sub

Private Sub txtInteresSolicitado_LostFocus()
On Error GoTo vError
If CCur(IIf((txtInteresSolicitado = ""), 0, txtInteresSolicitado)) >= 0 And CCur(IIf((txtPlazoSolicitado = ""), 0, txtPlazoSolicitado)) > 0 _
    And CCur(IIf((medMontoSolicitado = ""), 0, medMontoSolicitado)) > 0 Then
  txtCuotaSolicitada = fxCalcula_Cuota(CCur(medMontoSolicitado), CCur(txtPlazoSolicitado), CCur(txtInteresSolicitado))
End If
vError:
End Sub

Private Sub txtNombre_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF4 Then Call sbBusqueda(3)
End Sub

Private Sub txtOperacion_Change()
 Call LimpiaDatos
  With tlbPrincipal.Buttons
   .Item(2).Enabled = False
   .Item(3).Enabled = False
   .Item(4).Enabled = False
 End With
End Sub

Private Sub txtOperacion_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then Call sbCargaOperacion
If KeyCode = vbKeyF4 Then Call sbBusqueda(0)
End Sub

Private Sub txtPlazoSolicitado_Change()
On Error GoTo vError
If CCur(IIf((txtInteresSolicitado = ""), 0, txtInteresSolicitado)) >= 0 And CCur(IIf((txtPlazoSolicitado = ""), 0, txtPlazoSolicitado)) > 0 _
    And CCur(IIf((medMontoSolicitado = ""), 0, medMontoSolicitado)) > 0 Then
  txtCuotaSolicitada = fxCalcula_Cuota(CCur(medMontoSolicitado), CCur(txtPlazoSolicitado), CCur(txtInteresSolicitado))
End If

vError:
End Sub


Private Sub txtPlazoSolicitado_KeyDown(KeyCode As Integer, Shift As Integer)
Dim X As Double

 If KeyCode = vbKeyReturn And txtPlazoSolicitado.Text <> "" Then
   
   X = fxCatalogoRangoPlz(txtCodigo, txtPlazoSolicitado, fxCodigoCbo(cboDestino))
   
   If X > 0 Then
     txtInteresSolicitado.Text = X
   End If
 End If
End Sub

Private Sub txtPlazoSolicitado_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then txtInteresSolicitado.SetFocus
End Sub

Private Sub medMontoAprobado_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then txtPlazoAprobado.SetFocus
End Sub

Private Sub txtInteresAprobado_Change()
On Error Resume Next
If CCur(IIf((txtInteresAprobado = ""), 0, txtInteresAprobado)) >= 0 And CCur(IIf((txtPlazoAprobado = ""), 0, txtPlazoAprobado)) > 0 _
    And CCur(IIf((medMontoAprobado = ""), 0, medMontoAprobado)) > 0 Then
  txtCuotaAprobada = fxCalcula_Cuota(CCur(medMontoAprobado), CCur(txtPlazoAprobado), CCur(txtInteresAprobado))
End If
End Sub

Private Sub medMontoAprobado_Change()
On Error Resume Next
If CCur(IIf((txtInteresAprobado = ""), 0, txtInteresAprobado)) >= 0 And CCur(IIf((txtPlazoAprobado = ""), 0, txtPlazoAprobado)) > 0 _
    And CCur(IIf((medMontoAprobado = ""), 0, medMontoAprobado)) > 0 Then
 txtCuotaAprobada = fxCalcula_Cuota(CCur(medMontoAprobado), CCur(txtPlazoAprobado), CCur(txtInteresAprobado))
End If
End Sub

Private Sub txtInteresAprobado_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then dtpFechaRes.SetFocus
End Sub

Private Sub txtPlazoAprobado_Change()
On Error GoTo vError
If CCur(IIf((txtInteresAprobado = ""), 0, txtInteresAprobado)) >= 0 And CCur(IIf((txtPlazoAprobado = ""), 0, txtPlazoAprobado)) > 0 _
    And CCur(IIf((medMontoSolicitado = ""), 0, medMontoSolicitado)) > 0 Then
  txtCuotaAprobada = fxCalcula_Cuota(CCur(medMontoAprobado), CCur(txtPlazoAprobado), CCur(txtInteresAprobado))
End If
vError:
End Sub


Private Sub txtPlazoAprobado_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyReturn Then txtInteresAprobado.SetFocus
End Sub

