VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.Controls.v22.1.0.ocx"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#22.1#0"; "Codejock.ShortcutBar.v22.1.0.ocx"
Begin VB.Form frmMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H80000003&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Menú"
   ClientHeight    =   8280
   ClientLeft      =   120
   ClientTop       =   390
   ClientWidth     =   8415
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MouseIcon       =   "frmMenu.frx":0000
   ScaleHeight     =   8280
   ScaleWidth      =   8415
   ShowInTaskbar   =   0   'False
   Begin XtremeSuiteControls.PushButton btnRefresh 
      Height          =   312
      Left            =   8040
      TabIndex        =   16
      Top             =   120
      Width           =   360
      _Version        =   1441793
      _ExtentX        =   635
      _ExtentY        =   556
      _StockProps     =   79
      FlatStyle       =   -1  'True
      UseVisualStyle  =   -1  'True
      Appearance      =   17
      Picture         =   "frmMenu.frx":6852
   End
   Begin XtremeSuiteControls.GroupBox fraOpciones 
      Height          =   2175
      Left            =   1920
      TabIndex        =   4
      Top             =   3000
      Visible         =   0   'False
      Width           =   3375
      _Version        =   1441793
      _ExtentX        =   5953
      _ExtentY        =   3836
      _StockProps     =   79
      Caption         =   "Favoritos"
      ForeColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   16
      BorderStyle     =   1
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   735
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Agregar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmMenu.frx":6F52
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   735
         Index           =   2
         Left            =   2160
         TabIndex        =   6
         Top             =   1320
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Cerrar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmMenu.frx":7672
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.PushButton btnFavoritos 
         Height          =   735
         Index           =   1
         Left            =   1200
         TabIndex        =   7
         Top             =   1320
         Width           =   975
         _Version        =   1441793
         _ExtentX        =   1720
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Eliminar"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         UseVisualStyle  =   -1  'True
         Appearance      =   17
         Picture         =   "frmMenu.frx":7CB0
         TextImageRelation=   1
      End
      Begin XtremeSuiteControls.Label lblOpcion 
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   3135
         _Version        =   1441793
         _ExtentX        =   5530
         _ExtentY        =   1085
         _StockProps     =   79
         Caption         =   "Label1"
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Transparent     =   -1  'True
         WordWrap        =   -1  'True
      End
   End
   Begin XtremeShortcutBar.ShortcutBar wndShortcutBar 
      Height          =   7536
      Left            =   0
      TabIndex        =   10
      Top             =   480
      Width           =   3732
      _Version        =   1441793
      _ExtentX        =   6583
      _ExtentY        =   13293
      _StockProps     =   64
      VisualTheme     =   3
   End
   Begin MSComctlLib.ImageList imgMenuNodos 
      Left            =   720
      Top             =   480
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
            Picture         =   "frmMenu.frx":83C6
            Key             =   "Aplicacion"
            Object.Tag             =   "Aplicacion"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":8B36
            Key             =   "Aplicaciones"
            Object.Tag             =   "Aplicaciones"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":92A7
            Key             =   "Carpeta"
            Object.Tag             =   "Carpeta"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":9A4A
            Key             =   "Reportes"
            Object.Tag             =   "Reportes"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":A206
            Key             =   "Configuracion"
            Object.Tag             =   "Configuracion"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":AB6B
            Key             =   "Documento"
            Object.Tag             =   "Documento"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":B324
            Key             =   "Estadistica"
            Object.Tag             =   "Estadistica"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":BA0A
            Key             =   "Reloj"
            Object.Tag             =   "Reloj"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":C409
            Key             =   "Ayuda"
            Object.Tag             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":CE54
            Key             =   "Root"
            Object.Tag             =   "Root"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":D62D
            Key             =   "Explorer"
            Object.Tag             =   "Explorer"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":13E8F
            Key             =   "Favoritos"
            Object.Tag             =   "Favoritos"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1461F
            Key             =   "Usuario"
            Object.Tag             =   "Usuario"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":14F82
            Key             =   "Opciones"
            Object.Tag             =   "Opciones"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":15504
            Key             =   "Calendario"
            Object.Tag             =   "Calendario"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":15BEC
            Key             =   "Buscar"
            Object.Tag             =   "Buscar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":16162
            Key             =   "Libros"
            Object.Tag             =   "Libros"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1C9C4
            Key             =   "Direccion"
            Object.Tag             =   "Direccion"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":23226
            Key             =   "Cajas"
            Object.Tag             =   "Cajas"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":29A88
            Key             =   "Identificacion"
            Object.Tag             =   "Identificacion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2A1E0
            Key             =   "Agenda"
            Object.Tag             =   "Agenda"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2A760
            Key             =   "Histograma"
            Object.Tag             =   "Histograma"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2ACEB
            Key             =   "Administrador"
            Object.Tag             =   "Administrador"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":3154D
            Key             =   "Analisis"
            Object.Tag             =   "Analisis"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":31F1B
            Key             =   "Seguridad"
            Object.Tag             =   "Seguridad"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":32907
            Key             =   "Compras"
            Object.Tag             =   "Compras"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":330AF
            Key             =   "Caja Fuerte"
            Object.Tag             =   "Caja Fuerte"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":39911
            Key             =   "Dinero 1"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":40173
            Key             =   "Dinero 2"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":469D5
            Key             =   "Ajustes"
            Object.Tag             =   "Ajustes"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":47089
            Key             =   "Contacto"
            Object.Tag             =   "Contacto"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":471CC
            Key             =   "FastFoward"
            Object.Tag             =   "FastFoward"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":479D1
            Key             =   "Dinero 3"
            Object.Tag             =   "Dinero 3"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":47AFC
            Key             =   "Grafico"
            Object.Tag             =   "Grafico"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":47C35
            Key             =   "Favorito Add"
            Object.Tag             =   "Favorito Add"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":485A9
            Key             =   "Aprobacion"
            Object.Tag             =   "Aprobacion"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":48D87
            Key             =   "Lupa"
            Object.Tag             =   "Lupa"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":497A5
            Key             =   "Printer 2"
            Object.Tag             =   "Printer 2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":49D30
            Key             =   "Exportar"
            Object.Tag             =   "Exportar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgMenuLista 
      Left            =   1920
      Top             =   20
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   41
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":50592
            Key             =   "Root"
            Object.Tag             =   "Root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":50C88
            Key             =   "Reloj"
            Object.Tag             =   "Reloj"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5140F
            Key             =   "Ayuda"
            Object.Tag             =   "Ayuda"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":51A60
            Key             =   "Calendario"
            Object.Tag             =   "Calendario"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":521CC
            Key             =   "Dinero 2"
            Object.Tag             =   "Dinero 2"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":52970
            Key             =   "Contacto"
            Object.Tag             =   "Contacto"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":530A2
            Key             =   "Direccion"
            Object.Tag             =   "Direccion"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":53784
            Key             =   "Libros"
            Object.Tag             =   "Libros"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":53ED7
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5465B
            Key             =   "Printer 2"
            Object.Tag             =   "Printer 2"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":54D2C
            Key             =   "Exportar"
            Object.Tag             =   "Exportar"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":554E3
            Key             =   "Agenda"
            Object.Tag             =   "Agenda"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":55C09
            Key             =   "Lupa"
            Object.Tag             =   "Lupa"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":562DE
            Key             =   "Carpeta"
            Object.Tag             =   "Carpeta"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":56977
            Key             =   "Administrador"
            Object.Tag             =   "Administrador"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":56FD9
            Key             =   "Favorito Add"
            Object.Tag             =   "Favorito Add"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":57631
            Key             =   "Ajustes"
            Object.Tag             =   "Ajustes"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":57CFD
            Key             =   "Dinero 3"
            Object.Tag             =   "Dinero 3"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":583F1
            Key             =   "Documento"
            Object.Tag             =   "Documento"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":58BE6
            Key             =   "Dinero 1"
            Object.Tag             =   "Dinero 1"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":593D0
            Key             =   "Grafico"
            Object.Tag             =   "Grafico"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":59B09
            Key             =   "Seguridad"
            Object.Tag             =   "Seguridad"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5A14D
            Key             =   "Compras"
            Object.Tag             =   "Compras"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5A816
            Key             =   "Aplicacion"
            Object.Tag             =   "Aplicacion"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5AE34
            Key             =   "Aplicaciones"
            Object.Tag             =   "Aplicaciones"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5B50B
            Key             =   "Configuracion"
            Object.Tag             =   "Configuracion"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5BB83
            Key             =   "Estadistica"
            Object.Tag             =   "Estadistica"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5C245
            Key             =   "Analisis"
            Object.Tag             =   "Analisis"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5C8C3
            Key             =   "Explorer"
            Object.Tag             =   "Explorer"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5CFF0
            Key             =   "Aprobacion"
            Object.Tag             =   "Aprobacion"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5D6C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5DE3E
            Key             =   "Opciones"
            Object.Tag             =   "Opciones"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5E5DA
            Key             =   "Histograma"
            Object.Tag             =   "Histograma"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5EDC8
            Key             =   "Usuario"
            Object.Tag             =   "Usuario"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5F48C
            Key             =   "Identificacion"
            Object.Tag             =   "Identificacion"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":5FB94
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":602D0
            Key             =   "Caja Fuerte"
            Object.Tag             =   "Caja Fuerte"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":60A5F
            Key             =   "Buscar"
            Object.Tag             =   "Buscar"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":61156
            Key             =   "FastFoward"
            Object.Tag             =   "FastFoward"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":61821
            Key             =   "Cajas"
            Object.Tag             =   "Cajas"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":61F32
            Key             =   "Reportes"
            Object.Tag             =   "Reportes"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgProGrX 
      Left            =   1320
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":625A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":62B47
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":630C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":635D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":63B1F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00FF8080&
      BorderStyle     =   0  'None
      FillColor       =   &H00FF8080&
      Height          =   2160
      Left            =   3720
      ScaleHeight     =   940.557
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   72
   End
   Begin MSComctlLib.StatusBar StatusBarX 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   8025
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Enabled         =   0   'False
            Object.ToolTipText     =   "Usuario"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Bevel           =   0
            Object.ToolTipText     =   "Fecha"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   6068
            MinWidth        =   6068
            Object.ToolTipText     =   "Agencia / Oficina"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Bevel           =   0
            Enabled         =   0   'False
            Object.Width           =   776
            MinWidth        =   776
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   0
            Object.Width           =   1658
            MinWidth        =   1658
            TextSave        =   "12:25:p. m."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lsw 
      Height          =   6132
      Left            =   3840
      TabIndex        =   0
      Top             =   480
      Width           =   3252
      _ExtentX        =   5741
      _ExtentY        =   10821
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      PictureAlignment=   3
      _Version        =   393217
      Icons           =   "imgMenuLista"
      SmallIcons      =   "imgMenuLista"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   6068
      EndProperty
   End
   Begin MSComctlLib.TreeView vTreeMenu 
      Height          =   6135
      Index           =   0
      Left            =   5640
      TabIndex        =   2
      Top             =   480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   10821
      _Version        =   393217
      Indentation     =   648
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgMenuNodos"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeSuiteControls.FlatEdit txtBuscar 
      Height          =   312
      Left            =   3720
      TabIndex        =   9
      Top             =   120
      Width           =   4332
      _Version        =   1441793
      _ExtentX        =   7641
      _ExtentY        =   550
      _StockProps     =   77
      ForeColor       =   0
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Alignment       =   2
      Appearance      =   2
      UseVisualStyle  =   0   'False
   End
   Begin MSComctlLib.TreeView vTreeMenu 
      Height          =   6132
      Index           =   1
      Left            =   5640
      TabIndex        =   11
      Top             =   360
      Width           =   3372
      _ExtentX        =   5953
      _ExtentY        =   10821
      _Version        =   393217
      Indentation     =   648
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgMenuNodos"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView vTreeMenu 
      Height          =   6132
      Index           =   2
      Left            =   5520
      TabIndex        =   12
      Top             =   480
      Width           =   3372
      _ExtentX        =   5953
      _ExtentY        =   10821
      _Version        =   393217
      Indentation     =   648
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgMenuNodos"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView vTreeMenu 
      Height          =   6132
      Index           =   3
      Left            =   5400
      TabIndex        =   13
      Top             =   1320
      Width           =   3372
      _ExtentX        =   5953
      _ExtentY        =   10821
      _Version        =   393217
      Indentation     =   648
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgMenuNodos"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.TreeView vTreeMenu 
      Height          =   6132
      Index           =   4
      Left            =   5640
      TabIndex        =   14
      Top             =   360
      Width           =   3372
      _ExtentX        =   5953
      _ExtentY        =   10821
      _Version        =   393217
      Indentation     =   648
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgMenuNodos"
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutTitulo 
      Height          =   315
      Left            =   0
      TabIndex        =   15
      Top             =   120
      Width           =   3732
      _Version        =   1441793
      _ExtentX        =   6583
      _ExtentY        =   556
      _StockProps     =   14
      Caption         =   "Cuentas Corrientes"
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
   Begin VB.Image imgSplitter 
      Height          =   2388
      Left            =   3720
      MousePointer    =   9  'Size W E
      Top             =   480
      Width           =   156
   End
   Begin VB.Image imgMenuOption 
      Height          =   240
      Index           =   1
      Left            =   3480
      Picture         =   "frmMenu.frx":6403B
      Tag             =   "0"
      ToolTipText     =   "Tipo de Vista"
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim vPaso As Boolean, vNode As Node, mbMoving As Boolean, x As Boolean
Const sglSplitLimit = 0

Const SHORTCUT_CUENTAS = 300
Const SHORTCUT_RETAIL = 301
Const SHORTCUT_ADMINISTRATIVOS = 302
Const SHORTCUT_FINANCIEROS = 303
Const SHORTCUT_CONFIGURACION = 304

Const SHORTCUT_FOLDER_LIST = 305
Const SHORTCUT_SHORTCUTS = 306
Const SHORTCUT_JOURNAL = 307

Const SHORTCUT_SHOW_MORE = 308
Const SHORTCUT_SHOW_FEWER = 309


Private Sub SizeControls(x As Single)
    On Error Resume Next

    'set the width
    
    lsw.Left = x + 40
    lsw.Width = Me.Width - (lsw.Left + 120)
    
    wndShortcutBar.Width = lsw.Left - 80
    
    imgSplitter.Left = wndShortcutBar.Width + wndShortcutBar.Left + 40

    lsw.Left = wndShortcutBar.Width + wndShortcutBar.Left + 80
    
        
    txtBuscar.Visible = True
    
    lsw.top = wndShortcutBar.top
    imgSplitter.top = wndShortcutBar.top
    
    
    wndShortcutBar.Height = Me.Height - (StatusBarX.Height + 1120)
    
    lsw.Height = wndShortcutBar.Height
    imgSplitter.Height = wndShortcutBar.Height
    
    
    ShortcutTitulo.Width = wndShortcutBar.Width
    txtBuscar.Left = lsw.Left
    txtBuscar.Width = lsw.Width - (120 + btnRefresh.Width)
     
    'Posicion de los Iconos del Menu
    btnRefresh.Left = txtBuscar.Left + txtBuscar.Width + 20
    
    lsw.ColumnHeaders.Item(1).Width = lsw.Width - 20
     
End Sub


Private Sub btnFavoritos_Click(Index As Integer)
Dim pTrayIcon As XtremeSuiteControls.TrayIcon

Set pTrayIcon = frmContenedor.TrayIcon

Select Case Index
    Case 0
      Call sbFavoritosAdd(lblOpcion.Tag, "+")
              
        pTrayIcon.ShowBalloonTip 25, "ProGrX: Menú" _
                    , lblOpcion.Caption & " agregado a la Lista de Favoritos!" _
                    , xtpToolTipIconInfo
    Case 1
      Call sbFavoritosAdd(lblOpcion.Tag, "-")
        pTrayIcon.ShowBalloonTip 25, "ProGrX: Menú" _
                    , lblOpcion.Caption & " eliminado de la Lista de Favoritos!" _
                    , xtpToolTipIconInfo
    Case 2
End Select

fraOpciones.Visible = False
Call sbFavoritos

End Sub

Private Sub btnRefresh_Click()
    Call sbMenus
End Sub

Private Sub Form_Resize()
    On Error Resume Next
'    If Me.Width < 3000 Then Me.Width = 3000
'    SizeControls imgSplitter.Left - tlbMenu.Width
SizeControls picSplitter.Left

'    SizeControls vTreeMenu.Left + vTreeMenu.Width + 40
End Sub


Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub


Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    

    If mbMoving Then
'        picSplitter.Left = x + imgSplitter.Left
        
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
End Sub


Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    picSplitter.Left = imgSplitter.Left + x
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub

Private Sub sbCreaNodos(pTree As TreeView, vPadre As String, vTexto As String, vImagen As String, vExpand As Boolean _
            , Optional xkey As String = "N", Optional vBond As Boolean = False, Optional ptag As String = "")
Dim nodX As Node, vKey As String

On Error Resume Next

Set nodX = pTree.Nodes.Add(vPadre, tvwChild)
    nodX.Image = vImagen
    nodX.Text = vTexto
    nodX.Tag = ptag
    If xkey = "N" Then
        nodX.Key = vTexto & "0x0" & pTree.Nodes.Count & "ID"
    Else
        nodX.Key = xkey
    End If
    
    If vBond Then nodX.Bold = True
    
Set vNode = nodX

If vExpand Then
    Set nodX = pTree.Nodes.Add(xkey, tvwChild)
        nodX.Key = "F" & vTexto & "0x0" & pTree.Nodes.Count & "ID"
        nodX.Tag = nodX.Index
End If

End Sub

Private Sub sbMenuSub(Optional Index As Integer = 0, Optional pFiltro As String = "1,2,3,4,5,6,7,8,9,10,18")
Dim strSQL As String, rs As New ADODB.Recordset
Dim xNode As Node, lng As Long, i As Integer

Me.MousePointer = vbHourglass

vPaso = True

With vTreeMenu.Item(Index)
  .Nodes.Clear
  'Crear Root
  Set xNode = .Nodes.Add(, , "Main", "ProGrX", "Root")
  xNode.Bold = True
  
  strSQL = "select * from US_Menus" _
         & " where tipo = 'M' and dbo.fxSEG_MenuAccess(" & gPortal.Empresa_Id & ",'" & glogon.Usuario & "', Modulo, Formulario, Tipo) = 1" _
         & " and Modulo in(" & pFiltro & ")" _
         & " and nodo_padre is null order by  prioridad"
  Call OpenRecordSet(rs, strSQL, 1)
  Do While Not rs.EOF
   Call sbCreaNodos(vTreeMenu.Item(Index), "Main", Trim(rs!NODO_DESCRIPCION), rs!Icono, True, "0x0" & rs!menu_nodo, True)
    .Nodes(.Nodes.Count).Expanded = True
   rs.MoveNext
  Loop
  rs.Close

   xNode.Expanded = True

End With

vPaso = False

Me.MousePointer = vbDefault

End Sub

Private Sub sbMenus()

'Carga los Menus de todos las familias

Call sbMenuSub(0, "1,2,3,4,5,6,7,9,10,18")
Call sbMenuSub(1, "9,30,31,32,33,34,35")
Call sbMenuSub(2, "8,11,14,16,17,19,21,23,30,31,37,38,40")
Call sbMenuSub(3, "12,20,21,22,24,30,31,36")
Call sbMenuSub(4, "0")
           
End Sub

Private Sub sbFavoritos()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem

On Error GoTo vError

Me.MousePointer = vbHourglass

vPaso = True

With lsw.ListItems
  .Clear
  
  strSQL = "exec spSEG_MenuFavoritos " & gPortal.Empresa_Id & ",'" & glogon.Usuario & "'"
  
  Call OpenRecordSet(rs, strSQL, 1)
  Do While Not rs.EOF
    Set itmX = .Add(, ("0x0" & rs!menu_nodo), Trim(rs!NODO_DESCRIPCION), Trim(rs!Icono), Trim(rs!Icono))
        itmX.Tag = rs!menu_nodo
   rs.MoveNext
  Loop
  rs.Close
End With

Me.MousePointer = vbDefault
vPaso = False

Exit Sub

vError:
  Me.MousePointer = vbDefault
  vPaso = False
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub


Private Sub sbBuscar()
Dim strSQL As String, rs As New ADODB.Recordset
Dim itmX As ListItem


Me.MousePointer = vbHourglass
On Error Resume Next
vPaso = True

With lsw.ListItems
  .Clear
  
  strSQL = "select * From US_MENUS" _
         & " where dbo.fxSEG_MenuAccess(" & gPortal.Empresa_Id & ",'" & glogon.Usuario & "', Modulo, Formulario, Tipo) = 1" _
         & "   and NODO_DESCRIPCION like '%" & txtBuscar.Text & "%' and Tipo = 'A'"
  Call OpenRecordSet(rs, strSQL, 1)
  Do While Not rs.EOF
    Set itmX = .Add(, ("0x0" & rs!menu_nodo), Trim(rs!NODO_DESCRIPCION), Trim(rs!Icono), Trim(rs!Icono))
        itmX.Tag = rs!menu_nodo
   rs.MoveNext
  Loop
  rs.Close
End With

Me.MousePointer = vbDefault

vPaso = False

End Sub



Private Sub sbCreateShortcutBar()

    Dim Item As ShortcutBarItem ', ItemMain As ShortcutBarItem
 
     wndShortcutBar.Icons.LoadBitmap App.Path & "\Icons\Cuentas_24x24.png", SHORTCUT_CUENTAS, xtpImageNormal
     wndShortcutBar.Icons.LoadBitmap App.Path & "\Icons\Calculadora.png", SHORTCUT_RETAIL, xtpImageNormal
 
     wndShortcutBar.Icons.LoadBitmap App.Path & "\Icons\Administrativos_24x24.png", SHORTCUT_ADMINISTRATIVOS, xtpImageNormal
     wndShortcutBar.Icons.LoadBitmap App.Path & "\Icons\Financieros_24x24.png", SHORTCUT_FINANCIEROS, xtpImageNormal
     wndShortcutBar.Icons.LoadBitmap App.Path & "\Icons\Monitoreo_24x24.png", SHORTCUT_CONFIGURACION, xtpImageNormal
 
 
    
'    Set ItemMain = wndShortcutBar.AddItem(SHORTCUT_CUENTAS, "Cuentas Corrientes", vTreeMenu.Item(0).hWnd)
    Set Item = wndShortcutBar.AddItem(SHORTCUT_CUENTAS, "Cuentas Corrientes", vTreeMenu.Item(0).hWnd)
        Item.Selected = True
    Set Item = wndShortcutBar.AddItem(SHORTCUT_RETAIL, "Retail", vTreeMenu.Item(1).hWnd)
    Set Item = wndShortcutBar.AddItem(SHORTCUT_ADMINISTRATIVOS, "Administrativos", vTreeMenu.Item(2).hWnd)
    Set Item = wndShortcutBar.AddItem(SHORTCUT_FINANCIEROS, "Financieros", vTreeMenu.Item(3).hWnd)
    Set Item = wndShortcutBar.AddItem(SHORTCUT_CONFIGURACION, "Configuración", vTreeMenu.Item(4).hWnd)
       
       

'    wndShortcutBar.Selected = ItemMain

End Sub

Private Sub Form_Load()
Dim strSQL As String, rs As New ADODB.Recordset
Dim i As Integer

On Error GoTo vError


For i = 0 To vTreeMenu.Count - 1
    vTreeMenu.Item(i).Left = vTreeMenu.Item(i).Width * -1
Next i


strSQL = "select descripcion,dbo.MyGetdate() as Fecha from SIF_oficinas where cod_oficina = '" & GLOBALES.gOficinaTitular & "'"
Call OpenRecordSet(rs, strSQL)
If rs.EOF And rs.BOF Then
    StatusBarX.Panels(1).Text = glogon.Usuario
    StatusBarX.Panels(2).Text = Format(fxFechaServidor, "dd/mm/yyyy")
    StatusBarX.Panels(3).Text = GLOBALES.gOficinaTitular
Else
    StatusBarX.Panels(1).Text = glogon.Usuario
    StatusBarX.Panels(2).Text = Format(rs!fecha, "dd/mm/yyyy")
    StatusBarX.Panels(3).Text = Trim(rs!Descripcion)
End If
rs.Close

Me.Width = 7890
Me.Height = MDIPrincipal.Height - 2000
Me.Caption = "Menú: ProGrX"

Me.BackColor = RGB(70, 111, 178)



'Me.Caption = "Menú                  [ " & Trim(GLOBALES.gstrNombreEmpresa) & " ]"

'Set lsw.Picture = fxImagen_Leer("select Logo from SIF_Empresa", "Logo")

Call sbCreateShortcutBar
Call sbMenus
Call sbFavoritos



Exit Sub

vError:
 MsgBox fxSys_Error_Handler(Err.Description), vbCritical

End Sub





Private Sub lsw_DblClick()
If lsw.ListItems.Count <= 0 Then Exit Sub

Call sbSIFMenuOptionClick(lsw.SelectedItem.Tag)

End Sub


Private Sub lsw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If lsw.ListItems.Count <= 0 Then Exit Sub

If Button = 2 Then
   lblOpcion.Caption = lsw.SelectedItem
   lblOpcion.Tag = lsw.SelectedItem.Tag
   
   fraOpciones.Visible = True
   fraOpciones.Left = x
   fraOpciones.top = y
End If

End Sub

Private Sub sbFavoritosAdd(pNodo As Long, Optional pOpcion As String = "+")
Dim strSQL As String

On Error GoTo vError

strSQL = "exec spSEG_MenuFavoritosAdd " & gPortal.Empresa_Id & ",'" & glogon.Usuario & "'," & pNodo & ",'" & pOpcion & "'"
Call ConectionExecute(strSQL, 1)

Exit Sub

vError:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
  

End Sub



'Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
'Dim i As Integer
'
'If Button.Value = tbrUnpressed Then
'  Button.Value = tbrPressed
'  Exit Sub
'End If
'
'For i = 1 To tlbMenu.Buttons.Count
'    If tlbMenu.Buttons.Item(i).Key <> Button.Key Then
'       tlbMenu.Buttons.Item(i).Value = tbrUnpressed
'    End If
'Next i
'
'Call sbMenus
'
'End Sub

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Or KeyCode = vbKeyTab Then
  Call sbBuscar
End If
End Sub


Private Sub vTreeMenu_DblClick(Index As Integer)
On Error GoTo vError
If vTreeMenu.Item(Index).Nodes.Count <= 0 Then Exit Sub
If vTreeMenu.Item(Index).SelectedItem.Key = "Main" Then Exit Sub
Call sbSIFMenuOptionClick(Mid(vTreeMenu.Item(Index).SelectedItem.Key, 4, Len(vTreeMenu.Item(Index).SelectedItem.Key)))
vError:

End Sub

Private Sub vTreeMenu_Expand(Index As Integer, ByVal Node As MSComctlLib.Node)
Dim rs As New ADODB.Recordset, strSQL As String
Dim xNode As Node, lng As Long


If vPaso Then Exit Sub

If Node.Key = "Main" Then Exit Sub
If Node.Tag = "Expanded" Then Exit Sub

On Error GoTo vError

Me.MousePointer = vbHourglass

Node.Tag = "Expanded"
If Node.Index > 1 Then vTreeMenu.Item(Index).Nodes.Remove Node.Child.Index

With vTreeMenu.Item(Index)
  
  strSQL = "select * from US_Menus where nodo_padre = " & Mid(Node.Key, 4, Len(Node.Key)) _
        & " and dbo.fxSEG_MenuAccess(" & gPortal.Empresa_Id & ",'" & glogon.Usuario & "', Modulo, Formulario, Tipo) = 1" _
        & " order by prioridad"
        
  Call OpenRecordSet(rs, strSQL, 1)
  Do While Not rs.EOF
   Call sbCreaNodos(vTreeMenu.Item(Index), Node.Key, Trim(rs!NODO_DESCRIPCION), rs!Icono, IIf((rs!Tipo = "A"), False, True), "0x0" & rs!menu_nodo, False)
    .Nodes(.Nodes.Count).Expanded = True
   rs.MoveNext
  Loop
  rs.Close

End With


Me.MousePointer = vbDefault

Exit Sub

vError:
  Me.MousePointer = vbDefault
  MsgBox fxSys_Error_Handler(Err.Description), vbExclamation
End Sub

Private Sub vTreeMenu_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
On Error GoTo vError

If vTreeMenu.Item(Index).Nodes.Count <= 0 Then Exit Sub
If vTreeMenu.Item(Index).SelectedItem.Key = "Main" Then Exit Sub

If Button = 2 Then
   lblOpcion.Caption = vTreeMenu.Item(Index).SelectedItem
   lblOpcion.Tag = Mid(vTreeMenu.Item(Index).SelectedItem.Key, 4, Len(vTreeMenu.Item(Index).SelectedItem.Key))
   
   fraOpciones.Visible = True
   fraOpciones.Left = x
   fraOpciones.top = y
End If


vError:

End Sub

Private Sub wndShortcutBar_SelectedChanged(ByVal Item As XtremeShortcutBar.IShortcutBarItem)
    ShortcutTitulo.Caption = Item.Caption
End Sub
