VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUS_Menus 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Administración de Menús de Sistemas"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   11730
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgMenuLista 
      Left            =   7800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   38
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":0000
            Key             =   "Aplicacion"
            Object.Tag             =   "Aplicacion"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":01E5
            Key             =   "Dinero 1"
            Object.Tag             =   "Dinero 1"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":6A47
            Key             =   "Dinero 2"
            Object.Tag             =   "Dinero 2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":D2A9
            Key             =   "Aplicaciones"
            Object.Tag             =   "Aplicaciones"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":D47C
            Key             =   "Carpeta"
            Object.Tag             =   "Carpeta"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":D67E
            Key             =   "Reportes"
            Object.Tag             =   "Reportes"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":D898
            Key             =   "Configuracion"
            Object.Tag             =   "Configuracion"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":DA6B
            Key             =   "Documento"
            Object.Tag             =   "Documento"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":DC76
            Key             =   "Estadistica"
            Object.Tag             =   "Estadistica"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":DE5B
            Key             =   "Reloj"
            Object.Tag             =   "Reloj"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":E0BA
            Key             =   "Ayuda"
            Object.Tag             =   "Ayuda"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":E2E3
            Key             =   "Root"
            Object.Tag             =   "Root"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":E4F1
            Key             =   "Explorer"
            Object.Tag             =   "Explorer"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":E6AB
            Key             =   "Favoritos"
            Object.Tag             =   "Favoritos"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":E88C
            Key             =   "Usuario"
            Object.Tag             =   "Usuario"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":EA43
            Key             =   "Opciones"
            Object.Tag             =   "Opciones"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":EC4B
            Key             =   "Calendario"
            Object.Tag             =   "Calendario"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":EE7E
            Key             =   "Buscar"
            Object.Tag             =   "Buscar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":F07B
            Key             =   "Libros"
            Object.Tag             =   "Libros"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":158DD
            Key             =   "Direccion"
            Object.Tag             =   "Direccion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":1C13F
            Key             =   "Identificacion"
            Object.Tag             =   "Identificacion"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":229A1
            Key             =   "Cajas"
            Object.Tag             =   "Cajas"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":29203
            Key             =   "Agenda"
            Object.Tag             =   "Agenda"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":2FA65
            Key             =   "Histograma"
            Object.Tag             =   "Histograma"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":362C7
            Key             =   "Administrador"
            Object.Tag             =   "Administrador"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":3CB29
            Key             =   "Analisis"
            Object.Tag             =   "Analisis"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":4338B
            Key             =   "Seguridad"
            Object.Tag             =   "Seguridad"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":49BED
            Key             =   "Caja Fuerte"
            Object.Tag             =   "Caja Fuerte"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":5044F
            Key             =   "Ajustes"
            Object.Tag             =   "Ajustes"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":50687
            Key             =   "Contacto"
            Object.Tag             =   "Contacto"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":508CC
            Key             =   "FastFoward"
            Object.Tag             =   "FastFoward"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":50B48
            Key             =   "Dinero 3"
            Object.Tag             =   "Dinero 3"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":50D79
            Key             =   "Grafico"
            Object.Tag             =   "Grafico"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":50FA2
            Key             =   "Favorito Add"
            Object.Tag             =   "Favorito Add"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":511A4
            Key             =   "Aprobacion"
            Object.Tag             =   "Aprobacion"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":513C7
            Key             =   "Lupa"
            Object.Tag             =   "Lupa"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":515B2
            Key             =   "Printer 2"
            Object.Tag             =   "Printer 2"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":5180B
            Key             =   "Exportar"
            Object.Tag             =   "Exportar"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtFormulario 
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
      Left            =   7200
      TabIndex        =   24
      Top             =   5880
      Width           =   4455
   End
   Begin VB.CheckBox chkModal 
      Appearance      =   0  'Flat
      Caption         =   "Pantalla Modal"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   9600
      TabIndex        =   22
      Top             =   5520
      Width           =   1455
   End
   Begin MSComctlLib.ImageList imgMenuNodos_01 
      Left            =   7080
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":51A23
            Key             =   "Aplicacion"
            Object.Tag             =   "Aplicacion"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":51B41
            Key             =   "Aplicaciones"
            Object.Tag             =   "Aplicaciones"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":51C4F
            Key             =   "Carpeta"
            Object.Tag             =   "Carpeta"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":51D6B
            Key             =   "Reportes"
            Object.Tag             =   "Reportes"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":51E85
            Key             =   "Configuracion"
            Object.Tag             =   "Configuracion"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":51F85
            Key             =   "Documento"
            Object.Tag             =   "Documento"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":52089
            Key             =   "Estadistica"
            Object.Tag             =   "Estadistica"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":521A2
            Key             =   "Reloj"
            Object.Tag             =   "Reloj"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":522C2
            Key             =   "Ayuda"
            Object.Tag             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":523FB
            Key             =   "Root"
            Object.Tag             =   "Root"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":5251B
            Key             =   "Explorer"
            Object.Tag             =   "Explorer"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":5260F
            Key             =   "Favoritos"
            Object.Tag             =   "Favoritos"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":5271D
            Key             =   "Usuario"
            Object.Tag             =   "Usuario"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":52824
            Key             =   "Opciones"
            Object.Tag             =   "Opciones"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":52943
            Key             =   "Calendario"
            Object.Tag             =   "Calendario"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":52A70
            Key             =   "Buscar"
            Object.Tag             =   "Buscar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":52B79
            Key             =   "Libros"
            Object.Tag             =   "Libros"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":593DB
            Key             =   "Direccion"
            Object.Tag             =   "Direccion"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":5FC3D
            Key             =   "Identificacion"
            Object.Tag             =   "Identificacion"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":6649F
            Key             =   "Cajas"
            Object.Tag             =   "Cajas"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":6CD01
            Key             =   "Agenda"
            Object.Tag             =   "Agenda"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":73563
            Key             =   "Histograma"
            Object.Tag             =   "Histograma"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":79DC5
            Key             =   "Administrador"
            Object.Tag             =   "Administrador"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":80627
            Key             =   "Analisis"
            Object.Tag             =   "Analisis"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":86E89
            Key             =   "Seguridad"
            Object.Tag             =   "Seguridad"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":8D6EB
            Key             =   "Compras"
            Object.Tag             =   "Compras"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":93F4D
            Key             =   "Caja Fuerte"
            Object.Tag             =   "Caja Fuerte"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":9A7AF
            Key             =   "Dinero 1"
            Object.Tag             =   "Dinero 1"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A1011
            Key             =   "Dinero 2"
            Object.Tag             =   "Dinero 2"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A7873
            Key             =   "Ajustes"
            Object.Tag             =   "Ajustes"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A79A8
            Key             =   "Contacto"
            Object.Tag             =   "Contacto"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A7AEB
            Key             =   "FastFoward"
            Object.Tag             =   "FastFoward"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A7C32
            Key             =   "Dinero 3"
            Object.Tag             =   "Dinero 3"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A7D5D
            Key             =   "Grafico"
            Object.Tag             =   "Grafico"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A7E96
            Key             =   "Favorito Add"
            Object.Tag             =   "Favorito Add"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A7FC3
            Key             =   "Aprobacion"
            Object.Tag             =   "Aprobacion"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A80F8
            Key             =   "Lupa"
            Object.Tag             =   "Lupa"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A8216
            Key             =   "Printer 2"
            Object.Tag             =   "Printer 2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A8356
            Key             =   "Exportar"
            Object.Tag             =   "Exportar"
         EndProperty
      EndProperty
   End
   Begin VB.ComboBox cboModo 
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
      Height          =   330
      ItemData        =   "frmUS_Menus.frx":A848C
      Left            =   7200
      List            =   "frmUS_Menus.frx":A8496
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   5520
      Width           =   1695
   End
   Begin VB.ComboBox cboIconos 
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
      Height          =   330
      ItemData        =   "frmUS_Menus.frx":A84AD
      Left            =   9600
      List            =   "frmUS_Menus.frx":A84AF
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   5160
      Width           =   2055
   End
   Begin VB.ComboBox cboTipo 
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
      Height          =   330
      ItemData        =   "frmUS_Menus.frx":A84B1
      Left            =   7200
      List            =   "frmUS_Menus.frx":A84BE
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   5160
      Width           =   1695
   End
   Begin VB.TextBox txtPrioridad 
      Alignment       =   2  'Center
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
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0"
      Top             =   6960
      Width           =   855
   End
   Begin VB.TextBox txtClase 
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
      Left            =   7200
      TabIndex        =   16
      Top             =   6240
      Width           =   1695
   End
   Begin VB.TextBox txtClaseIdCall 
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
      Left            =   9600
      TabIndex        =   15
      Top             =   6240
      Width           =   2055
   End
   Begin VB.TextBox txtClaseDescripcion 
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
      Left            =   7200
      TabIndex        =   13
      Top             =   6600
      Width           =   4455
   End
   Begin VB.TextBox txtNodo 
      Alignment       =   2  'Center
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
      Height          =   315
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   4200
      Width           =   3975
   End
   Begin MSComctlLib.ImageList imgExp01 
      Left            =   3720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A84E2
            Key             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A85D6
            Key             =   "Formularios"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A86F4
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A87ED
            Key             =   "Modulos"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView vTree 
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   11880
      _Version        =   393217
      Indentation     =   648
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgExp01"
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
   Begin MSComctlLib.TreeView vTreeMenu 
      Height          =   3495
      Left            =   5760
      TabIndex        =   1
      Top             =   480
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   6165
      _Version        =   393217
      Indentation     =   648
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgMenuNodos"
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
   Begin MSComctlLib.Toolbar tlb 
      Height          =   264
      Left            =   9600
      TabIndex        =   11
      Top             =   4680
      Width           =   2112
      _ExtentX        =   3731
      _ExtentY        =   476
      ButtonWidth     =   487
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
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
      EndProperty
   End
   Begin MSComCtl2.FlatScrollBar FlatScrollBar 
      Height          =   255
      Left            =   8160
      TabIndex        =   21
      Top             =   6960
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   450
      _Version        =   393216
      Arrows          =   65536
      Orientation     =   1638401
   End
   Begin MSComctlLib.ImageList imgMenuLista_02 
      Left            =   9720
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A890B
            Key             =   "Aplicacion"
            Object.Tag             =   "Aplicacion"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A8F29
            Key             =   "Aplicaciones"
            Object.Tag             =   "Aplicaciones"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A9600
            Key             =   "Reportes"
            Object.Tag             =   "Reportes"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":A9CD1
            Key             =   "Configuracion"
            Object.Tag             =   "Configuracion"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":AA349
            Key             =   "Carpeta"
            Object.Tag             =   "Carpeta"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":AA9E2
            Key             =   "Documento"
            Object.Tag             =   "Documento"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":AB1D7
            Key             =   "Estadistica"
            Object.Tag             =   "Estadistica"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":AB899
            Key             =   "Reloj"
            Object.Tag             =   "Reloj"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":AC020
            Key             =   "Ayuda"
            Object.Tag             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":AC671
            Key             =   "Root"
            Object.Tag             =   "Root"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":ACD67
            Key             =   "Explorer"
            Object.Tag             =   "Explorer"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":AD51E
            Key             =   "Favoritos"
            Object.Tag             =   "Favoritos"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":ADB76
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":AE287
            Key             =   "Administrador"
            Object.Tag             =   "Administrador"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":AE8E9
            Key             =   "Calendario"
            Object.Tag             =   "Calendario"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":AF055
            Key             =   "Opciones"
            Object.Tag             =   "Opciones"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":AF7F1
            Key             =   "FastFoward"
            Object.Tag             =   "FastFoward"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":AFEBC
            Key             =   "Buscar"
            Object.Tag             =   "Buscar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B059C
            Key             =   "Favorito Add"
            Object.Tag             =   "Favorito Add"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B0BEB
            Key             =   "Usuario"
            Object.Tag             =   "Usuario"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B12AF
            Key             =   "Identificacion"
            Object.Tag             =   "Identificacion"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B19B7
            Key             =   "Analisis"
            Object.Tag             =   "Analisis"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B20F0
            Key             =   "Compras"
            Object.Tag             =   "Compras"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B27B9
            Key             =   "Cajas"
            Object.Tag             =   "Cajas"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B2F0C
            Key             =   "Agenda"
            Object.Tag             =   "Agenda"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B36B0
            Key             =   "Histograma"
            Object.Tag             =   "Histograma"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B3E9E
            Key             =   "Dinero 1"
            Object.Tag             =   "Dinero 1"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B451C
            Key             =   "Dinero 2"
            Object.Tag             =   "Dinero 2"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B4D06
            Key             =   "Dinero 3"
            Object.Tag             =   "Dinero 3"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B53FA
            Key             =   "Seguridad"
            Object.Tag             =   "Seguridad"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B5A3E
            Key             =   "Contacto"
            Object.Tag             =   "Contacto"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B6170
            Key             =   "Aprobacion"
            Object.Tag             =   "Aprobacion"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B6842
            Key             =   "Lupa"
            Object.Tag             =   "Lupa"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B6EDB
            Key             =   "Libros"
            Object.Tag             =   "Libros"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B7568
            Key             =   "Direccion"
            Object.Tag             =   "Direccion"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B7C9A
            Key             =   "Caja Fuerte"
            Object.Tag             =   "Caja Fuerte"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B8429
            Key             =   "Grafico"
            Object.Tag             =   "Grafico"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B8B52
            Key             =   "Ajustes"
            Object.Tag             =   "Ajustes"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B921E
            Key             =   "Printer 2"
            Object.Tag             =   "Printer 2"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":B98EF
            Key             =   "Exportar"
            Object.Tag             =   "Exportar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMenuNodos 
      Left            =   9000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":BA06B
            Key             =   "Aplicacion"
            Object.Tag             =   "Aplicacion"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":BA7DB
            Key             =   "Aplicaciones"
            Object.Tag             =   "Aplicaciones"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":BAF4C
            Key             =   "Carpeta"
            Object.Tag             =   "Carpeta"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":BB6EF
            Key             =   "Reportes"
            Object.Tag             =   "Reportes"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":BBEAB
            Key             =   "Configuracion"
            Object.Tag             =   "Configuracion"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":BC810
            Key             =   "Documento"
            Object.Tag             =   "Documento"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":BCFC9
            Key             =   "Estadistica"
            Object.Tag             =   "Estadistica"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":BD6AF
            Key             =   "Reloj"
            Object.Tag             =   "Reloj"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":BE0AE
            Key             =   "Ayuda"
            Object.Tag             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":BEAF9
            Key             =   "Root"
            Object.Tag             =   "Root"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":BF2D2
            Key             =   "Explorer"
            Object.Tag             =   "Explorer"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":C5B34
            Key             =   "Favoritos"
            Object.Tag             =   "Favoritos"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":C62C4
            Key             =   "Usuario"
            Object.Tag             =   "Usuario"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":C6C27
            Key             =   "Opciones"
            Object.Tag             =   "Opciones"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":C71A9
            Key             =   "Calendario"
            Object.Tag             =   "Calendario"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":C7891
            Key             =   "Buscar"
            Object.Tag             =   "Buscar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":C7E07
            Key             =   "Libros"
            Object.Tag             =   "Libros"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":CE669
            Key             =   "Direccion"
            Object.Tag             =   "Direccion"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":D4ECB
            Key             =   "Cajas"
            Object.Tag             =   "Cajas"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":DB72D
            Key             =   "Identificacion"
            Object.Tag             =   "Identificacion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":DBE85
            Key             =   "Agenda"
            Object.Tag             =   "Agenda"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":DC405
            Key             =   "Histograma"
            Object.Tag             =   "Histograma"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":DC990
            Key             =   "Administrador"
            Object.Tag             =   "Administrador"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":E31F2
            Key             =   "Analisis"
            Object.Tag             =   "Analisis"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":E3BC0
            Key             =   "Seguridad"
            Object.Tag             =   "Seguridad"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":E45AC
            Key             =   "Compras"
            Object.Tag             =   "Compras"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":E4D54
            Key             =   "Caja Fuerte"
            Object.Tag             =   "Caja Fuerte"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":EB5B6
            Key             =   "Dinero 1"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":F1E18
            Key             =   "Dinero 2"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":F867A
            Key             =   "Ajustes"
            Object.Tag             =   "Ajustes"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":F8D2E
            Key             =   "Contacto"
            Object.Tag             =   "Contacto"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":F8E71
            Key             =   "FastFoward"
            Object.Tag             =   "FastFoward"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":F9676
            Key             =   "Dinero 3"
            Object.Tag             =   "Dinero 3"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":F97A1
            Key             =   "Grafico"
            Object.Tag             =   "Grafico"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":F98DA
            Key             =   "Favorito Add"
            Object.Tag             =   "Favorito Add"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":FA24E
            Key             =   "Aprobacion"
            Object.Tag             =   "Aprobacion"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":FAA2C
            Key             =   "Lupa"
            Object.Tag             =   "Lupa"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":FB44A
            Key             =   "Printer 2"
            Object.Tag             =   "Printer 2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUS_Menus.frx":FB9D5
            Key             =   "Exportar"
            Object.Tag             =   "Exportar"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgRefrescar 
      Height          =   240
      Index           =   2
      Left            =   4920
      Picture         =   "frmUS_Menus.frx":102237
      ToolTipText     =   "Encoger"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgRefrescar 
      Height          =   240
      Index           =   1
      Left            =   4560
      Picture         =   "frmUS_Menus.frx":102346
      ToolTipText     =   "Refresca Menus"
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   10800
      TabIndex        =   30
      Top             =   6960
      Width           =   135
   End
   Begin VB.Label lblNodoPadre 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11040
      TabIndex        =   29
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label lblNodo 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   10200
      TabIndex        =   28
      Top             =   6960
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "Nodo / Padre"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   10
      Left            =   8760
      TabIndex        =   27
      Top             =   6960
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Módulo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   9
      Left            =   5640
      TabIndex        =   26
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Formulario"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   5640
      TabIndex        =   25
      Top             =   5880
      Width           =   1455
   End
   Begin VB.Label lblModulo 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   23
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Image imgModulos 
      Height          =   240
      Left            =   11280
      Picture         =   "frmUS_Menus.frx":10245F
      ToolTipText     =   "Exportar Autmáticamente Módulos"
      Top             =   4200
      Width           =   240
   End
   Begin VB.Image imgRefrescar 
      Height          =   240
      Index           =   0
      Left            =   11280
      Picture         =   "frmUS_Menus.frx":102564
      ToolTipText     =   "Refresca Menus"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image imgIcono 
      Height          =   255
      Left            =   11280
      ToolTipText     =   "Muestra del Icono a Utilizar"
      Top             =   5520
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Icono"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   7
      Left            =   9000
      TabIndex        =   14
      Top             =   5160
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Prioridad / Orden"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   6
      Left            =   5640
      TabIndex        =   10
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Descripción"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   5640
      TabIndex        =   9
      Top             =   6600
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Cls. Id."
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   9000
      TabIndex        =   8
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Modo de Acceso"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   5640
      TabIndex        =   7
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Clase"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   5640
      TabIndex        =   6
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5640
      TabIndex        =   5
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Nodo Actual"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   5640
      TabIndex        =   4
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Menus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   1
      Left            =   5640
      TabIndex        =   3
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Módulos y Formularios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "frmUS_Menus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vScroll As Boolean, vEdita As Boolean
Dim vNode As Node, vNodoActual As Node



Private Function fxMenuIcono(pEtiqueta As String) As Integer
Dim i As Integer


pEtiqueta = UCase(Trim(pEtiqueta))

For i = 1 To imgMenuNodos.ListImages.Count
  If UCase(imgMenuNodos.ListImages(i).Tag) = pEtiqueta Then Exit For
Next i

fxMenuIcono = i

End Function


Private Sub cboIconos_Click()
On Error GoTo vError

imgIcono.Picture = imgMenuNodos.ListImages.Item(fxMenuIcono(cboIconos.Text)).Picture

vError:
End Sub



Private Sub sbCreaNodos(pTree As TreeView, vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean _
            , Optional xKey As String = "N", Optional vBond As Boolean = False, Optional pTag As String = "")
Dim nodX As Node, vKey As String

On Error Resume Next

Set nodX = pTree.Nodes.Add(vPadre, tvwChild)
    nodX.Image = vImagen
    nodX.Text = vTexto
    nodX.Tag = pTag
    If xKey = "N" Then
        nodX.Key = vTexto & "0x0" & pTree.Nodes.Count & "ID"
    Else
        nodX.Key = xKey
    End If
    
    If vBond Then nodX.Bold = True
    
Set vNode = nodX

If vExpand Then
    Set nodX = pTree.Nodes.Add(xKey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & pTree.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If

End Sub


Private Sub sbMenus()
Dim strSQL As String, rs As New ADODB.Recordset
Dim xNode As Node, lng As Long


Me.MousePointer = vbHourglass

vPaso = True

With vTreeMenu
  .Nodes.Clear
  'Crear Root
  Set xNode = .Nodes.Add(, , "Main", "PGX", "Root")
  xNode.Bold = True
  
  strSQL = "select * from US_Menus where tipo = 'M' and nodo_padre is null order by prioridad"
  Call OpenRecordSet(rs, strSQL)
  Do While Not rs.EOF
   Call sbCreaNodos(vTreeMenu, "Main", Trim(rs!NODO_DESCRIPCION), rs!Icono, True, "0x0" & rs!menu_nodo, True)
    .Nodes(.Nodes.Count).Expanded = True
   rs.MoveNext
  Loop
  rs.Close

   xNode.Expanded = True

End With

Me.MousePointer = vbDefault

vPaso = False

End Sub


Private Sub sbCargaInicial()
Dim strSQL As String, rs As New ADODB.Recordset
Dim xNode As Node, lng As Long


Me.MousePointer = vbHourglass

With vTree
  .Nodes.Clear
  'Crear Root
  Set xNode = .Nodes.Add(, , "US", "PGX", "Root")
  xNode.Bold = True
  
  strSQL = "select * from us_modulos order by modulo"
  Call OpenRecordSet(rs, strSQL, 1)
  Do While Not rs.EOF
   Call sbCreaNodos(vTree, "US", Trim(rs!Nombre), "Modulos", False, "0x0" & rs!Modulo & "M", True)
   rs.MoveNext
  .Nodes(.Nodes.Count).Expanded = True
  Loop
  rs.Close

  strSQL = "select *,dbo.fxSEG_OpcionAsignada(Formulario,0) as 'Existe' from US_formularios order by formulario"
  Call OpenRecordSet(rs, strSQL, 1)
  Do While Not rs.EOF
   Call sbCreaNodos(vTree, "0x0" & rs!Modulo & "M", Trim(rs!Descripcion), IIf(rs!Existe = 1, "Check", "Formularios") _
                , False, "0x0" & Trim(rs!Formulario) & "F", False, rs!Formulario)
   rs.MoveNext
  .Nodes(.Nodes.Count).Expanded = True
  Loop
  rs.Close
  
   xNode.Expanded = True

End With


Me.MousePointer = vbDefault

End Sub



Private Sub FlatScrollBar_Change()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vPrioridad As Long, NodoTemp As Node
Dim vKey1 As String, vKey2 As String
Dim vImage As String, vText As String, vTag As String

On Error GoTo vError

If lblNodo.Caption = "0" Then
   MsgBox "Debe Guardar el Nodo antes de Priorizar ? ", vbExclamation
   Exit Sub
End If


If vScroll Then
    
    If FlatScrollBar.Value = 1 Then
       vPrioridad = CLng(txtPrioridad.Text) + 1
        'Nodo Actual con esa Prioridad
        strSQL = "select menu_nodo  from us_menus where nodo_padre = " & lblNodoPadre.Caption _
               & " and Prioridad = " & vPrioridad
        Call OpenRecordSet(rs, strSQL, 1)
        If Not rs.EOF And Not rs.BOF Then
            strSQL = "update us_menus set prioridad = " & vPrioridad - 1 & " where menu_nodo = " & rs!menu_nodo
            Call ConectionExecute(strSQL, 1)
        End If
        rs.Close
    
        Set NodoTemp = vTreeMenu.Nodes.Item(vNode.Index).Next
    
    Else
       vPrioridad = CLng(txtPrioridad.Text) - 1
        'Nodo Actual con esa Prioridad
        strSQL = "select menu_nodo from us_menus where nodo_padre = " & lblNodoPadre.Caption _
               & " and Prioridad = " & vPrioridad
        Call OpenRecordSet(rs, strSQL)
        If Not rs.EOF And Not rs.BOF Then
            strSQL = "update us_menus set prioridad = " & vPrioridad + 1 & " where menu_nodo = " & rs!menu_nodo
            Call ConectionExecute(strSQL, 1)
        End If
        rs.Close
    
        Set NodoTemp = vTreeMenu.Nodes.Item(vNode.Index).Previous
    End If
    
    strSQL = "update us_menus set prioridad = " & vPrioridad & " where menu_nodo = " & lblNodo.Caption
    Call ConectionExecute(strSQL, 1)
    
    txtPrioridad.Text = vPrioridad
    
    vKey1 = vNode.Key
    vKey2 = NodoTemp.Key
   
    vImage = vNode.Image
    vText = vNode.Text
    vTag = vNode.Tag

   
   
    With vTreeMenu.Nodes.Item(vNode.Index)
       .Key = "0x0000000000000001"
       .Text = NodoTemp.Text
       .Tag = NodoTemp.Tag
       .Image = NodoTemp.Image
    End With
    
    With vTreeMenu.Nodes.Item(NodoTemp.Index)
       .Key = "0x0000000000000002"
       .Text = vText
       .Tag = vTag
       .Image = vImage
    End With
    
    
    vTreeMenu.Nodes.Item(vNode.Index).Key = vKey2
    vTreeMenu.Nodes.Item(NodoTemp.Index).Key = vKey1
    
End If

vScroll = False
FlatScrollBar.Value = 0
vScroll = True

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub Form_Load()
vModulo = 13
 
 
 vScroll = False
  FlatScrollBar.Value = 0
 vScroll = True


 vEdita = True
 Call sbToolBarIconos(tlb)
 Call sbToolBar(tlb, "Nuevo")
 
 
 'Llena Combos
  With cboIconos
     .Clear
     .AddItem "Aplicacion"
     .AddItem "Aplicaciones"
     .AddItem "Carpeta"
     .AddItem "Reportes"
     .AddItem "Configuracion"
     .AddItem "Documento"
     .AddItem "Estadistica"
     .AddItem "Reloj"
     .AddItem "Ayuda"
     .AddItem "Root"
     .AddItem "Explorer"
     .AddItem "Favoritos"
     .AddItem "Usuario"
     .AddItem "Opciones"
     .AddItem "Calendario"
     .AddItem "Buscar"
     .AddItem "Libros"
     .AddItem "Direccion"
     .AddItem "Identificacion"
     .AddItem "Cajas"
     .AddItem "Agenda"
     .AddItem "Histograma"
     .AddItem "Administrador"
     .AddItem "Analisis"
     .AddItem "Seguridad"
     .AddItem "Compras"
     .AddItem "Caja Fuerte"
     .AddItem "Dinero 1"
     .AddItem "Dinero 2"
     .AddItem "Dinero 3"
     .AddItem "Ajustes"
     .AddItem "Contacto"
     .AddItem "FastFoward"
     .AddItem "Grafico"
     .AddItem "Favorito Add"
     .AddItem "Aprobacion"
     .AddItem "Lupa"
     .AddItem "Printer 2"
     .AddItem "Exportar"

     .Text = "Aplicacion"
  End With
 
 Call sbCargaInicial
 Call sbMenus
 
 Call sbLimpia
   
 
 Call Formularios(Me)
 Call RefrescaTags(Me)

End Sub


Private Sub sbLimpia()


cboIconos.Text = "Aplicacion"
cboTipo.Text = "Acceso Directo"
cboModo.Text = "Formulario"


txtFormulario.Text = ""
txtClase.Text = ""
txtClaseIdCall.Text = "0"
txtClaseDescripcion.Text = ""

chkModal.Value = vbUnchecked

lblNodo.Caption = 0


End Sub


Private Function fxMenuNodoExist() As Integer
Dim strSQL As String, rs As New ADODB.Recordset


End Function


Private Function fxMenuNodoPrioridad(NodoPadre As Long) As Integer
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer


If NodoPadre = 0 Then
    strSQL = "select isnull(max(prioridad),1000) + 1 as 'Prioridad'" _
           & " from us_menus where Nodo_Padre is null"
Else
    strSQL = "select isnull(max(prioridad),1000) + 1 as 'Prioridad'" _
           & " from us_menus where Nodo_Padre = " & NodoPadre
End If
       
Call OpenRecordSet(rs, strSQL, 1)
i = rs!prioridad
rs.Close

fxMenuNodoPrioridad = i

End Function



Private Sub imgModulos_Click()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vMenuNodo As Long, vMenuPrioridad As Long

Me.MousePointer = vbHourglass


strSQL = "select isnull(max(menu_nodo),0) + 1 as MenuNodo from us_Menus"
Call OpenRecordSet(rs, strSQL)
  vMenuNodo = rs!MenuNodo
rs.Close

'vMenuPrioridad = fxMenuNodoPrioridad(0)

strSQL = "select * from US_modulos" _
       & " where modulo not in(select modulo from us_menus where tipo = 'M') order by modulo"
Call OpenRecordSet(rs, strSQL)


Do While Not rs.EOF
 strSQL = "insert us_menus(menu_nodo,modulo,nodo_padre,nodo_descripcion,tipo,icono,modo,modal" _
        & ",formulario,accesos_dll_id,accesos_dll_cls,prioridad) values(" & vMenuNodo & "," & rs!Modulo _
        & ",Null,'" & rs!Nombre & "','M','Aplicacion','M',0,'',0,''," _
        & rs!Modulo * 100 & ")"
        
 Call ConectionExecute(strSQL, 1)
 
 vMenuNodo = vMenuNodo + 1
 vMenuPrioridad = vMenuPrioridad + 1
 
 rs.MoveNext
Loop
rs.Close

Me.MousePointer = vbDefault

Call sbMenus

End Sub



Private Sub imgRefrescar_Click(Index As Integer)
Dim lng As Long
 
Select Case Index
 Case 0 'Menus
     Call sbMenus
 Case 1 'Opciones
     Call sbCargaInicial
 Case 2 'Encoger
 
        With vTree.Nodes
         For lng = 1 To .Count
          If Right(.Item(lng).Key, 1) = "M" Then
            .Item(lng).Expanded = IIf(.Item(lng).Expanded, False, True)
          End If
         Next lng
        End With
End Select
End Sub


Private Sub sbGuardar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim vNodoNuevo As Long, vNodoPadre As Long, vNodePadreFill As String
Dim vPrioridad As Long

On Error GoTo vError

If lblNodo.Caption = 0 Then
  
 strSQL = "select isnull(max(menu_nodo),0) + 1 as Nodo from us_menus"
 Call OpenRecordSet(rs, strSQL, 1)
   vNodoNuevo = rs!nodo
 rs.Close
 
 
 If Mid(cboTipo.Text, 1, 1) = "M" And CLng(lblModulo.Caption) = 0 Then
    MsgBox "Indique una Opción de algún módulo primero, el modulo no puede ser 0", vbExclamation
    Exit Sub
 End If
 
 If Mid(cboTipo.Text, 1, 1) = "M" Then
     vNodoPadre = "0"
     vPrioridad = 1000 + CLng(lblModulo.Caption)
 Else
     vNodoPadre = lblNodoPadre.Caption
     vPrioridad = fxMenuNodoPrioridad(vNodoPadre)
 End If
 
 
 
 If vNodoPadre = 0 Then
    vNodePadreFill = " Null "
 Else
    vNodePadreFill = CStr(vNodoPadre)
 End If
 
 strSQL = "insert us_menus(menu_nodo,modulo,nodo_padre,nodo_descripcion,tipo,icono,modo,modal" _
        & ",formulario,accesos_dll_id,accesos_dll_cls,prioridad) values(" & vNodoNuevo & "," & lblModulo.Caption _
        & "," & vNodePadreFill & ",'" & txtClaseDescripcion.Text & "','" & Mid(cboTipo.Text, 1, 1) & "','" _
        & cboIconos.Text & "','" & Mid(cboModo.Text, 1, 1) & "'," & chkModal.Value & ",'" & txtFormulario.Text & "'," & txtClaseIdCall.Text _
        & ",'" & txtClase.Text & "'," & vPrioridad & ")"
 Call ConectionExecute(strSQL, 1)


   Call sbCreaNodos(vTreeMenu, "0x0" & vNodoPadre, Trim(txtClaseDescripcion.Text), cboIconos.Text, IIf((Mid(cboTipo.Text, 1, 1) = "A"), False, True), "0x0" & vNodoNuevo, False)
    vTreeMenu.Nodes(vTreeMenu.Nodes.Count).Expanded = True


Else
  strSQL = "update us_menus set nodo_descripcion = '" & txtClaseDescripcion.Text & "', Tipo = '" & Mid(cboTipo.Text, 1, 1) _
         & "', Modo = '" & Mid(cboModo.Text, 1, 1) & "', Modal = " & chkModal.Value & ", Formulario = '" & txtFormulario.Text _
         & "', Icono = '" & cboIconos.Text & "', accesos_dll_id = " & txtClaseIdCall.Text & ",accesos_dll_cls = '" _
         & txtClase.Text & "',prioridad = " & txtPrioridad.Text _
         & " where menu_nodo = " & lblNodo.Caption
  Call ConectionExecute(strSQL, 1)
  
  With vTreeMenu.SelectedItem
      .Image = cboIconos.Text
      .Text = txtClaseDescripcion.Text
   End With

End If


Call sbToolBar(tlb, "Activo")


If vNodoActual Is Nothing Then
Else
    vNodoActual.Image = "Check"
End If

Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub

Private Sub sbBorrar()
Dim strSQL As String

On Error GoTo vError

If lblNodo.Caption <> "0" Then
  
  strSQL = "delete us_menus_usos where menu_nodo in(select menu_nodo from" _
         & " us_menus where nodo_padre = " & lblNodo.Caption & ")"
  Call ConectionExecute(strSQL, 1)
  
  strSQL = "delete us_menus_usos where menu_nodo = " & lblNodo.Caption
  Call ConectionExecute(strSQL, 1)
  
  strSQL = "delete us_menus where nodo_padre = " & lblNodo.Caption
  Call ConectionExecute(strSQL, 1)
  
  strSQL = "delete us_menus where menu_nodo = " & lblNodo.Caption
  Call ConectionExecute(strSQL, 1)
  
  vTreeMenu.Nodes.Remove (vNode.Index)
End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical


End Sub

Private Sub tlb_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
  Case "nuevo"
    Call sbToolBar(tlb, "Edicion")
    
    If Mid(cboTipo.Text, 1, 1) <> "A" Then
       lblNodoPadre.Caption = lblNodo.Caption
    End If
    
    Call sbLimpia
    
  Case "editar"
    Call sbToolBar(tlb, "Edicion")
  Case "guardar"
    Call sbGuardar
  Case "borrar"
    Call sbBorrar
  Case "deshacer"
    Call sbToolBar(tlb, "Activo")

End Select
End Sub


Private Sub vTree_DblClick()
On Error GoTo vError

Call sbFormsCall(vNodoActual.Tag, , , , True, Me)

vError:
End Sub

Private Sub vTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError
 
'El Tag de los nodos fue cargado con el nombre del formulario

Set vNodoActual = Node

If Node.Tag <> "" Then
 If Node.Image = "Check" Then
   strSQL = "select menu_nodo from us_menus where formulario = '" & Node.Tag & "'"
   Call OpenRecordSet(rs, strSQL, 1)
   
   If Not rs.EOF And Not rs.BOF Then
       Call sbNodoConsulta(rs!menu_nodo)
   End If
   rs.Close
   
 Else
   strSQL = "select * from us_formularios where formulario = '" & Node.Tag & "'"
   Call OpenRecordSet(rs, strSQL, 1)
   
   If Not rs.EOF And Not rs.BOF Then
        Call sbToolBar(tlb, "Edicion")
   
        cboIconos.Text = "Aplicacion"
        cboTipo.Text = "Acceso Directo"
        cboModo.Text = "Formulario"
        txtFormulario = Trim(rs!Formulario)
        txtClase.Text = ""
        txtClaseIdCall.Text = 0
        txtClaseDescripcion.Text = Trim(rs!Descripcion)
        
        lblModulo.Caption = rs!Modulo
        
        chkModal.Value = vbUnchecked
        txtPrioridad.Text = "1001"
        If lblNodoPadre.Caption = 0 Then lblNodoPadre.Caption = lblNodo.Caption
        lblNodo.Caption = 0
   
   
   End If
   rs.Close
 
 End If

End If

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub

Private Function fxNodoIndice(pKey As String) As String
pKey = Mid(pKey, 4, Len(pKey))
pKey = Mid(pKey, 1, Len(pKey) - 1)
fxNodoIndice = pKey
End Function

Private Sub sbNodoConsulta(pKey As String)
Dim strSQL As String, rs As New ADODB.Recordset

On Error GoTo vError

If vNode.Key = "Main" Then Exit Sub

strSQL = "Select * from US_Menus where Menu_Nodo = " & pKey
Call OpenRecordSet(rs, strSQL)


cboIconos.Text = Trim(rs!Icono)
Select Case rs!Tipo
  Case "A"
      cboTipo.Text = "Acceso Directo"
      Call sbToolBar(tlb, "Edicion")
  
  Case "F"
      cboTipo.Text = "Folder"
      Call sbToolBar(tlb, "Activo")
  
  Case "M"
      cboTipo.Text = "Módulo"
      Call sbToolBar(tlb, "Activo")
End Select

Select Case rs!Modo
  Case "F"
      cboModo.Text = "Formulario"
  Case "C"
      cboModo.Text = "Clase"
End Select

txtFormulario = Trim(rs!Formulario)
txtClase.Text = Trim(rs!ACCESOS_DLL_CLS)
txtClaseIdCall.Text = Trim(rs!ACCESOS_DLL_ID)
txtClaseDescripcion.Text = Trim(rs!NODO_DESCRIPCION)

chkModal.Value = rs!Modal
txtPrioridad.Text = rs!prioridad

lblNodo.Caption = rs!menu_nodo
lblNodoPadre.Caption = IIf(IsNull(rs!nodo_padre), 0, rs!nodo_padre)
lblModulo.Caption = rs!Modulo

rs.Close

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  Call sbLimpia
End Sub


Private Sub vTreeMenu_AfterLabelEdit(Cancel As Integer, NewString As String)
Dim strSQL As String


strSQL = "update us_menus set nodo_descripcion = '" & NewString & "' where menu_nodo = " & Mid(vNode.Key, 4, Len(vNode.Key))
Call ConectionExecute(strSQL)

End Sub


Private Sub vTreeMenu_DblClick()
On Error GoTo vError


Call sbSIFMenuOptionClick(Mid(vNode.Key, 4, Len(vNode.Key)))

vError:

End Sub

Private Sub vTreeMenu_Expand(ByVal Node As MSComctlLib.Node)
Dim rs As New ADODB.Recordset, strSQL As String
Dim xNode As Node, lng As Long


If vPaso Then Exit Sub
If Node.Key = "Main" Then Exit Sub
If Node.Tag = "Expanded" Then Exit Sub

Me.MousePointer = vbHourglass

Node.Tag = "Expanded"
If Node.Index > 1 Then vTreeMenu.Nodes.Remove Node.Child.Index

With vTreeMenu
  
  strSQL = "select * from US_Menus where nodo_padre = " & Mid(Node.Key, 4, Len(Node.Key)) & " order by prioridad"
  Call OpenRecordSet(rs, strSQL, 1)
  
  Do While Not rs.EOF
   Call sbCreaNodos(vTreeMenu, Node.Key, Trim(rs!NODO_DESCRIPCION), rs!Icono, IIf((rs!Tipo = "A"), False, True), "0x0" & rs!menu_nodo, False)
    .Nodes(.Nodes.Count).Expanded = True
   rs.MoveNext
  Loop
  rs.Close

'   Node.Expanded = True

End With


Me.MousePointer = vbDefault



End Sub

Private Sub vTreeMenu_NodeClick(ByVal Node As MSComctlLib.Node)

txtNodo.Text = Node.FullPath
txtNodo.Tag = Node.Key

Set vNode = Node

Call sbNodoConsulta(Mid(Node.Key, 4, Len(Node.Key)))

End Sub
