VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "ComCt332.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#19.1#0"; "Codejock.CommandBars.v19.1.0.ocx"
Begin VB.MDIForm MDIPrincipal 
   BackColor       =   &H80000003&
   ClientHeight    =   10260
   ClientLeft      =   132
   ClientTop       =   780
   ClientWidth     =   15624
   HelpContextID   =   9010
   Icon            =   "MDIPrincipal.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIPrincipal.frx":071A
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer_Load 
      Interval        =   10
      Left            =   960
      Top             =   3000
   End
   Begin ComCtl3.CoolBar CoolBarContabilidad 
      Align           =   1  'Align Top
      Height          =   408
      Left            =   0
      TabIndex        =   3
      Top             =   408
      Visible         =   0   'False
      Width           =   15624
      _ExtentX        =   27559
      _ExtentY        =   720
      BandCount       =   4
      _CBWidth        =   15624
      _CBHeight       =   408
      _Version        =   "6.7.9816"
      Child1          =   "tlbPeriodo"
      MinHeight1      =   312
      Width1          =   1296
      NewRow1         =   0   'False
      MinWidth2       =   6000
      MinHeight2      =   360
      Width2          =   6000
      NewRow2         =   0   'False
      Child3          =   "tlbCierre"
      MinWidth3       =   1500
      MinHeight3      =   312
      Width3          =   1308
      NewRow3         =   0   'False
      Child4          =   "tlbEmpresa"
      MinHeight4      =   312
      NewRow4         =   0   'False
      BandStyle4      =   1
      Begin VB.TextBox lblPeriodo 
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   210
         Left            =   2685
         Locked          =   -1  'True
         TabIndex        =   9
         Text            =   "Periodo"
         Top             =   90
         Visible         =   0   'False
         Width           =   3585
      End
      Begin VB.TextBox txtMes 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   40
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.TextBox txtAnio 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1920
         TabIndex        =   7
         Top             =   40
         Visible         =   0   'False
         Width           =   615
      End
      Begin MSComctlLib.Toolbar tlbPeriodo 
         Height          =   312
         Left            =   132
         TabIndex        =   6
         Top             =   48
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   550
         ButtonWidth     =   1799
         ButtonHeight    =   550
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Periodos"
               ImageIndex      =   2
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbCierre 
         Height          =   312
         Left            =   7632
         TabIndex        =   5
         Top             =   48
         Width           =   7896
         _ExtentX        =   13928
         _ExtentY        =   550
         ButtonWidth     =   1736
         ButtonHeight    =   550
         Style           =   1
         TextAlignment   =   1
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cierres"
               Key             =   "cierres"
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
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Enabled         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Revisión"
               Key             =   "Revision"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Balance"
                     Text            =   "Revisión del Balance"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Asientos"
                     Text            =   "Verificación de Asientos"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbEmpresa 
         Height          =   312
         Left            =   15600
         TabIndex        =   4
         Top             =   48
         Width           =   24
         _ExtentX        =   42
         _ExtentY        =   550
         ButtonWidth     =   910
         ButtonHeight    =   466
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgContabilidad"
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
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   120
      Top             =   2280
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":2F5EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":35E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":3C6B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":42F12
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":49774
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":4FFD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":56838
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerSalir 
      Left            =   600
      Top             =   3000
   End
   Begin ComCtl3.CoolBar CoolBarX 
      Align           =   1  'Align Top
      Height          =   408
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15624
      _ExtentX        =   27559
      _ExtentY        =   720
      BandCount       =   5
      FixedOrder      =   -1  'True
      _CBWidth        =   15624
      _CBHeight       =   408
      _Version        =   "6.7.9816"
      Child1          =   "tlbMenu"
      MinWidth1       =   2208
      MinHeight1      =   264
      Width1          =   2208
      UseCoolbarColors1=   0   'False
      NewRow1         =   0   'False
      MinWidth2       =   1992
      MinHeight2      =   360
      Width2          =   1992
      NewRow2         =   0   'False
      Child3          =   "tlbFavoritos"
      MinWidth3       =   3060
      MinHeight3      =   264
      Width3          =   3060
      NewRow3         =   0   'False
      MinHeight4      =   360
      NewRow4         =   0   'False
      BandForeColor5  =   16777215
      BandBackColor5  =   13290186
      Caption5        =   "..."
      MinHeight5      =   336
      Width5          =   204
      UseCoolbarColors5=   0   'False
      NewRow5         =   0   'False
      BandStyle5      =   1
      BandEmbossShadow5=   -2147483636
      Begin MSComctlLib.Toolbar tlbFavoritos 
         Height          =   264
         Left            =   4584
         TabIndex        =   2
         Top             =   72
         Width           =   10608
         _ExtentX        =   18711
         _ExtentY        =   466
         ButtonWidth     =   487
         ButtonHeight    =   466
         Style           =   1
         ImageList       =   "imgMenuNodos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "EstadoCuenta"
               Object.ToolTipText     =   "Estado de Cuenta"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ConsultaCreditos"
               Object.ToolTipText     =   "Consulta General"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ConsultaExcedentes"
               Object.ToolTipText     =   "Consulta de Excedentes"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "CalculoOperacion"
               Object.ToolTipText     =   "Cálculo de la Operación"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ConsultaPersona"
               Object.ToolTipText     =   "Consulta de la Persona"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "ConsultaFondos"
               Object.ToolTipText     =   "Consulta de Fondos"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbMenu 
         Height          =   264
         Left            =   24
         TabIndex        =   1
         Top             =   72
         Width           =   2208
         _ExtentX        =   3895
         _ExtentY        =   466
         ButtonWidth     =   487
         ButtonHeight    =   466
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "imgMain"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   6
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Menu"
               Object.ToolTipText     =   "Menú Principal del Sistema"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Contabilidad"
               Object.ToolTipText     =   "Contabilidad: Explorador"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Activos"
               Object.ToolTipText     =   "Activos Fijos: Explorador"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Impresoras"
               Object.ToolTipText     =   "Configuración de Impresoras"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Marcas"
               Object.ToolTipText     =   "Control de Marcas"
               ImageIndex      =   6
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Marca"
                     Text            =   "Registrar Marca"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "MarcasDetalle"
                     Text            =   "Bitácora de Marcas"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Sep"
                     Text            =   "-"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Configuracion"
                     Text            =   "Configuración Horarios"
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "AsgUsuarios"
                     Text            =   "Asignación de Horarios"
                  EndProperty
               EndProperty
            EndProperty
         EndProperty
         Begin VB.Image Image3 
            Height          =   615
            Left            =   9240
            Picture         =   "MDIPrincipal.frx":5D09A
            Stretch         =   -1  'True
            Top             =   -720
            Width           =   375
         End
      End
   End
   Begin MSComctlLib.ImageList imgMenuNodos 
      Left            =   120
      Top             =   1680
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":5D3A4
            Key             =   "Aplicacion"
            Object.Tag             =   "Aplicacion"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":5DB14
            Key             =   "Aplicaciones"
            Object.Tag             =   "Aplicaciones"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":5E285
            Key             =   "Carpeta"
            Object.Tag             =   "Carpeta"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":5EA28
            Key             =   "Reportes"
            Object.Tag             =   "Reportes"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":5F1E4
            Key             =   "Configuracion"
            Object.Tag             =   "Configuracion"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":5FB49
            Key             =   "Documento"
            Object.Tag             =   "Documento"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":60302
            Key             =   "Estadistica"
            Object.Tag             =   "Estadistica"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":609E8
            Key             =   "Reloj"
            Object.Tag             =   "Reloj"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":613E7
            Key             =   "Ayuda"
            Object.Tag             =   "Ayuda"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":61E32
            Key             =   "Root"
            Object.Tag             =   "Root"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":6260B
            Key             =   "Explorer"
            Object.Tag             =   "Explorer"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":68E6D
            Key             =   "Favoritos"
            Object.Tag             =   "Favoritos"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":695FD
            Key             =   "Usuario"
            Object.Tag             =   "Usuario"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":69F60
            Key             =   "Opciones"
            Object.Tag             =   "Opciones"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":6A4E2
            Key             =   "Calendario"
            Object.Tag             =   "Calendario"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":6ABCA
            Key             =   "Buscar"
            Object.Tag             =   "Buscar"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":6B140
            Key             =   "Libros"
            Object.Tag             =   "Libros"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":719A2
            Key             =   "Direccion"
            Object.Tag             =   "Direccion"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":78204
            Key             =   "Cajas"
            Object.Tag             =   "Cajas"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":7EA66
            Key             =   "Identificacion"
            Object.Tag             =   "Identificacion"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":7F1BE
            Key             =   "Agenda"
            Object.Tag             =   "Agenda"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":7F73E
            Key             =   "Histograma"
            Object.Tag             =   "Histograma"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":7FCC9
            Key             =   "Administrador"
            Object.Tag             =   "Administrador"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":8652B
            Key             =   "Analisis"
            Object.Tag             =   "Analisis"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":86EF9
            Key             =   "Seguridad"
            Object.Tag             =   "Seguridad"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":878E5
            Key             =   "Compras"
            Object.Tag             =   "Compras"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":8808D
            Key             =   "Caja Fuerte"
            Object.Tag             =   "Caja Fuerte"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":8E8EF
            Key             =   "Dinero 1"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":95151
            Key             =   "Dinero 2"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":9B9B3
            Key             =   "Ajustes"
            Object.Tag             =   "Ajustes"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":9C067
            Key             =   "Contacto"
            Object.Tag             =   "Contacto"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":9C1AA
            Key             =   "FastFoward"
            Object.Tag             =   "FastFoward"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":9C9AF
            Key             =   "Dinero 3"
            Object.Tag             =   "Dinero 3"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":9CADA
            Key             =   "Grafico"
            Object.Tag             =   "Grafico"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":9CC13
            Key             =   "Favorito Add"
            Object.Tag             =   "Favorito Add"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":9D587
            Key             =   "Aprobacion"
            Object.Tag             =   "Aprobacion"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":9DD65
            Key             =   "Lupa"
            Object.Tag             =   "Lupa"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":9E783
            Key             =   "Printer 2"
            Object.Tag             =   "Printer 2"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":9ED0E
            Key             =   "Exportar"
            Object.Tag             =   "Exportar"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgContabilidad 
      Left            =   720
      Top             =   2280
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":A5570
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":ABDD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":B2634
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":B8E96
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":BF6F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":C5F5A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":CC7BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIPrincipal.frx":CCEE3
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   120
      Top             =   3000
      _Version        =   1245185
      _ExtentX        =   508
      _ExtentY        =   508
      _StockProps     =   0
      VisualTheme     =   5
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "&Archivo"
      Begin VB.Menu mnuSeguridad 
         Caption         =   "Seguridad"
         Begin VB.Menu mnuSeguridadSub 
            Caption         =   "Cambiar Contraseña"
            Index           =   0
         End
         Begin VB.Menu mnuSeguridadSub 
            Caption         =   "Actualiza Datos de Contacto"
            Index           =   1
         End
         Begin VB.Menu mnuSeguridadSub 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuSeguridadSub 
            Caption         =   "Cambiar de Tema"
            Index           =   3
         End
         Begin VB.Menu mnuSeguridadSub 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuSeguridadSub 
            Caption         =   "Bitácora"
            Index           =   5
         End
      End
      Begin VB.Menu mnuArchivoSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParametrosSistemaMenu 
         Caption         =   "Parámetros del Sistema"
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Datos de Empresa"
            Index           =   0
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Comunicados Generales"
            Index           =   1
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Encabezado y Pie de Página Estados"
            Index           =   3
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "-"
            Index           =   4
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Consulta de Cola de Asientos"
            Index           =   5
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "-"
            Index           =   6
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Oficinas"
            Index           =   7
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Oficinas : Metas de Colocación"
            Index           =   8
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "-"
            Index           =   9
         End
         Begin VB.Menu mnuParametrosSistema 
            Caption         =   "Variables Globales"
            Index           =   10
         End
         Begin VB.Menu mnuPE_Modo_1 
            Caption         =   "Planilla Especiales (Directas)"
         End
      End
      Begin VB.Menu mnuArchivoSeparador2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCambioEmpresa 
         Caption         =   "Cambio de Empresa"
      End
      Begin VB.Menu mnuArchivoSeparador21 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu mnuVerSub 
         Caption         =   "Ordenar por Iconos"
         Index           =   0
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "Ordenar en Cascada"
         Index           =   1
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "Ordenar Vertical"
         Index           =   2
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "Odenar Horizontal"
         Index           =   3
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "-"
         Index           =   4
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "Cerrar todas las ventanas"
         Index           =   5
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "Minimizar todas las ventanas"
         Index           =   6
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   "Restaurar todas la ventanas"
         Index           =   7
      End
      Begin VB.Menu mnuVerSub 
         Caption         =   ""
         Index           =   8
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "&?"
      Begin VB.Menu mnuAyudaContenido 
         Caption         =   "Contenido"
      End
      Begin VB.Menu mnuAyudaSoporteTecnico 
         Caption         =   "Soporte Técnico"
      End
      Begin VB.Menu mnuBarraHerramientas 
         Caption         =   "Barra de Herramientas"
      End
      Begin VB.Menu mnuAyudaSeparador1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAyudaAcercaDe 
         Caption         =   "Acerca De..."
      End
   End
   Begin VB.Menu mnuAcciones 
      Caption         =   "Acciones"
      Visible         =   0   'False
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Abonos"
         Index           =   0
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Anulación"
         Index           =   1
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Gestión de Cobro"
         Index           =   3
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Estado de la Operación"
         Index           =   4
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Nuevo Crédito"
         Index           =   6
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Trámites"
         Index           =   7
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "-"
         Index           =   8
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Historial"
         Index           =   9
         Visible         =   0   'False
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "-"
         Index           =   10
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Nuevo Análisis"
         Index           =   11
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Análisis ?"
         Index           =   12
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "-"
         Index           =   13
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Plan de Pagos"
         Index           =   14
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "-"
         Index           =   15
      End
      Begin VB.Menu mnuAccionesSub 
         Caption         =   "Cerrar"
         Index           =   16
      End
   End
   Begin VB.Menu mnuCxC 
      Caption         =   "CxC"
      Visible         =   0   'False
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Abonos"
         Index           =   0
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Anulación"
         Index           =   1
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Nueva Operación"
         Index           =   3
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Tramite"
         Index           =   4
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Movimientos"
         Index           =   6
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Plan de Pagos"
         Index           =   7
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "-"
         Index           =   8
      End
      Begin VB.Menu mnuCxCSub 
         Caption         =   "Cerrar"
         Index           =   9
      End
   End
   Begin VB.Menu mnuExplorerContable 
      Caption         =   "Explorador: Contable"
      Visible         =   0   'False
      Begin VB.Menu mnuCntAccionEditar 
         Caption         =   "&Editar"
      End
      Begin VB.Menu mnuCntAccionBorrar 
         Caption         =   "&Borrar"
      End
      Begin VB.Menu mnuCntSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCntAccionRefrescar 
         Caption         =   "Refrescar"
      End
      Begin VB.Menu mnuCntSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCntAccionesImprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu mnuCntSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCntAccionesMayorizar 
         Caption         =   "Mayorizar"
      End
   End
   Begin VB.Menu mnuActivosExplorador 
      Caption         =   "Explorardor: Activos"
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
      Begin VB.Menu mnuActivosAccionSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivosAccionDepreciacion 
         Caption         =   "Depreciación"
      End
      Begin VB.Menu mnuActivosAccionSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuActivosAccionActualizar 
         Caption         =   "Actualizar"
      End
      Begin VB.Menu mnuActivosAccionImprimir 
         Caption         =   "Imprimir"
      End
   End
End
Attribute VB_Name = "MDIPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public WithEvents statusBar As XtremeCommandBars.statusBar
Attribute statusBar.VB_VarHelpID = -1
Dim mLoad_Inicial As Boolean

Private Sub sbStatusBar_Load()
    Set statusBar = CommandBars.statusBar
    
    statusBar.Visible = True
    
    Dim Pane As StatusBarPane
    
'    Set Pane = StatusBar.AddPane(ID_INDICATOR_LOGO)
'    Pane.Text = "Codejock Software"
'    Pane.IconIndex = 100
'    'Pane.TextColor = vbGrayText
'    Pane.TextColor = RGB(64, 100, 176)
'    Pane.BackgroundColor = RGB(245, 245, 245)
'    Pane.Font.Bold = True
'    Pane.Width = 0 'Auto size
'    Pane.Visible = False

    Set Pane = statusBar.AddPane(0)
    Pane.Style = SBPS_STRETCH Or SBPS_NOBORDERS
    Pane.Text = "SystemLogic"
    Pane.Width = 0 ' Autro Size

    Pane.TextColor = vbWhite
    Pane.BackgroundColor = RGB(70, 111, 178)


    Set Pane = statusBar.AddPane(1)
    Pane.Style = SBPS_STRETCH Or SBPS_NOBORDERS
    Pane.Text = "Ready"
    Pane.Width = 0 ' Autro Size
    Pane.ToolTip = "Usuario:"
    Pane.TextColor = vbWhite
    Pane.BackgroundColor = RGB(70, 111, 178)

    Set Pane = statusBar.AddPane(2)
    Pane.Style = SBPS_STRETCH Or SBPS_NOBORDERS
    Pane.Text = "Ready"
    Pane.Width = 0 ' Autro Size
    Pane.ToolTip = "Fecha:"
    Pane.TextColor = vbWhite
    Pane.BackgroundColor = RGB(70, 111, 178)
    
    Set Pane = statusBar.AddPane(3)
    Pane.Style = SBPS_STRETCH Or SBPS_NOBORDERS
    Pane.Text = "Ready"
    Pane.Width = 0 ' Autro Size
    Pane.ToolTip = "Proceso"
    Pane.TextColor = vbWhite
    Pane.BackgroundColor = RGB(70, 111, 178)
   
    Set Pane = statusBar.AddPane(ID_INDICATOR_CAPS)
    Pane.Style = SBPS_NOBORDERS
    Pane.BeginGroup = True
    Pane.Dark = True
    Pane.TextColor = vbWhite
    Pane.BackgroundColor = RGB(70, 111, 178)
    
    Set Pane = statusBar.AddPane(ID_INDICATOR_NUM)
    Pane.Style = SBPS_NOBORDERS
    Pane.BeginGroup = False
    Pane.Dark = True
    Pane.TextColor = vbWhite
    Pane.BackgroundColor = RGB(70, 111, 178)
  
    Set Pane = statusBar.AddPane(ID_INDICATOR_SCRL)
    Pane.Style = SBPS_NOBORDERS
    Pane.BeginGroup = False
    Pane.Dark = True
    Pane.TextColor = vbWhite
    Pane.BackgroundColor = RGB(70, 111, 178)
    
    
    Me.BackColor = RGB(70, 111, 178)

End Sub

Private Sub MDIForm_Load()

mLoad_Inicial = True

If glogon.AppStatus = 1 Then
   Call sbFormsCall("frmCC_AppStatus", , , , False, Me)
End If

'Me.BackColor = RGB(36, 113, 163)
Me.BackColor = RGB(70, 111, 178)

CoolBarX.Bands.Item(4).BackColor = RGB(36, 113, 163)

'StatusBar.Panels(7).Text = glogon.ProGrX_Theme

Call sbStatusBar_Load

End Sub

Private Sub mnuCambioEmpresa_Click()
 Call Main_Reload
End Sub

'--- Menu Contextual: Contabilidad

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

'--- Fin: Menu Contextual: Contabilidad



'--- Menu Contextual: Activos Fijos

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

'---Fin de Menu Contextual: Activos Fijos

Private Sub MDIForm_Activate()
  
Me.Caption = App.ProductName & " [ " & App.Major & "." & App.Minor & "." & App.Revision & ".r" & GLOBALES.SysVersion & " ]"


txtAnio = gCntX_Parametros.PeriodoAnio
txtMes = gCntX_Parametros.PeriodoMes
tlbEmpresa.Buttons.Item(1).Caption = gCntX_Parametros.NombreEmpresa

End Sub



Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If UnloadMode = 0 Then
      Cancel = True
      TimerSalir.Interval = 10
   End If
End Sub


Public Sub mnuAccionesSub_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim x As clsEstudioCrd, vOperacion As Long
Dim vExpediente As String, vCajas As Boolean
Dim frmConsultaActiva As Form, frm As Form

On Error GoTo vError

vOperacion = 0
vExpediente = ""
vCajas = False

'Localiza el Frm de Consulta que se encuentre activo para utilizarlo de referencia
For Each frmConsultaActiva In Forms
  If (UCase(frmConsultaActiva.Name) = UCase("frmCR_ConsultaCreditos")) Then
    'lo localiza y sale (se supone que este procedimiento solo puede ser abierto desde esta pantalla
    Exit For
  End If
Next frmConsultaActiva

'Validacion

With frmConsultaActiva.vgCreditos
    .Sheet = .ActiveSheet
    .Row = .ActiveRow
    
    Select Case .Sheet
       Case 1 'Activos
          .col = 2
          If Index = 6 Or Index = 11 Then
            'Nada
           Else
              If Not IsNumeric(.CellTag) Then Exit Sub
              vOperacion = .CellTag
           End If
        
       Case 2, 3 'Cancelados y En Tramite
          .col = 2
          
          If Index = 6 Or Index = 11 Then
            'Nada
           Else
              If Not IsNumeric(.Text) Then Exit Sub
              vOperacion = .Text
           End If
       
       Case 4 'PreAnalisis
          .col = 2
          vExpediente = .Text
          .col = 7
          If IsNumeric(.Text) Then vOperacion = .Text
       
       Case 5 'Incobrables
          .col = 2
          If Not IsNumeric(.Text) Then Exit Sub
          vOperacion = .Text
      
    End Select
    

    Select Case Index
      Case 0 'Abonos
            If vOperacion = 0 Then Exit Sub
            .col = 6 'Saldo
            If CCur(.Text) = 0 Then Exit Sub
            
            vCajas = IIf((fxCajasParametros("01") = "S"), True, False)
                
                .col = 18 'Cuotas Morosas
                If CInt(.Text) = 0 Then
                  If vCajas Then
                        ModuloCajas.mRef_01 = vOperacion
                         
                        If GLOBALES.SysPlanPagos = 1 Then
                                 Call sbFormsCall("frmCajas_Crd_AbonosCtP", vbModal, 0, 0, False, Me, True)
                        Else
                                 Call sbFormsCall("frmCajas_Crd_AbonosStP", vbModal, 0, 0, False, Me, True)
                        End If
                  
                  Else
                        If GLOBALES.SysPlanPagos = 1 Then
                                 Call sbFormsCall("frmCR_AbonosNew", vbModal, 0, 0, False, Me, True)
                        Else
                                 Call sbFormsCall("frmCR_Abonos", vbModal, 0, 0, False, Me, True)
                        End If
                        
                        For Each frm In Forms
                          If (UCase(frm.Name) = UCase("frmCR_Abonos")) Or (UCase(frm.Name) = UCase("frmCR_AbonosNew")) Then
                            Call frm.sbConsultaExterna(vOperacion)
                            Exit For
                          End If
                        Next frm
                  End If
                Else 'Abonos en Mora
                   
                  If vCajas Then
                        ModuloCajas.mRef_01 = vOperacion
                         
                        If GLOBALES.SysPlanPagos = 1 Then
                                 Call sbFormsCall("frmCajas_Crd_AbonosCtP", vbModal, 0, 0, False, Me, True)
                        Else
                                 Call sbFormsCall("frmCajas_Crd_AbonosStP", vbModal, 0, 0, False, Me, True)
                        End If
                  
                  Else
                       If GLOBALES.SysPlanPagos = 1 Then
                                Call sbFormsCall("frmCR_AbonosNew")
                       Else
                                Call sbFormsCall("frmCR_CancelaMorosidad")
                       End If
                    
                        For Each frm In Forms
                          If (UCase(frm.Name) = UCase("frmCR_CancelaMorosidad")) Or (UCase(frm.Name) = UCase("frmCR_AbonosNew")) Then
                            Call frm.sbConsultaExterna(vOperacion)
                            Exit For
                          End If
                        Next frm
                   End If
                
                End If
    
      
      Case 1 'Anulacion de Abonos
            If vOperacion = 0 Then Exit Sub
                
                If GLOBALES.SysPlanPagos = 1 Then
                            Call sbFormsCall("frmCR_AnulaAbonosNew", 0, 0, 0, False, Me, True)
                Else
                            Call sbFormsCall("frmCR_AnulaAbonos", 0, 0, 0, False, Me, True)
                End If
      
                            For Each frm In Forms
                              If (UCase(frm.Name) = UCase("frmCR_AnulaAbonos")) Or (UCase(frm.Name) = UCase("frmCR_AnulaAbonosNew")) Then
                                Call frm.sbConsultaExterna(vOperacion)
                                Exit For
                              End If
                            Next frm
      
      Case 2 'Sep
      Case 3 'Gestion de Cobro
            If vOperacion = 0 Then Exit Sub
                
                Call sbFormsCall("frmCO_Principal")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCO_Principal") Then
                    Call frm.sbConsultaExterna(vOperacion)
                    Exit For
                  End If
                Next frm
      
      Case 4 'Movimientos de la Operacion
            
            If vOperacion = 0 Then Exit Sub
            
            Me.MousePointer = vbHourglass
            
            With frmContenedor.Crt
                .Reset
                .WindowShowPrintSetupBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowState = crptMaximized
                .WindowTitle = "Reportes del Módulo de Crédito"
                
                .Connect = glogon.ConectRPT
                
                If GLOBALES.SysPlanPagos = 0 Then
                    .ReportFileName = SIFGlobal.fxPathReportes("Cobro_AbonosOperacionFull.rpt")
                    .Formulas(0) = "Empresa='" & GLOBALES.gstrNombreEmpresa & "'"
                    .Formulas(1) = "SubTitulo='ABONOS ORDINARIOS/EXTRAORDINARIOS/MORATORIOS'"
                    .Formulas(2) = "Fecha='" & Format(fxFechaServidor, "dd/mm/yyyy") & "'"
                    .Formulas(3) = "Titulo='MOVIMIENTOS DE LA OPERACION'"
                    .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & vOperacion
                    
                    .SubreportToChange = "sbCorte"
                    .StoredProcParam(0) = vOperacion
                    .StoredProcParam(1) = Format(frmConsultaActiva.dtpCorte.Value, "yyyy/mm/dd")
                    
                    .SubreportToChange = "sbMovimientos"
                    
                    .StoredProcParam(0) = vOperacion
                    .StoredProcParam(1) = 1
                    
                Else
                     .ReportFileName = SIFGlobal.fxPathReportes("Credito_PlanPagosMov.rpt")
                    
                     .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy  hh:mm:ss") & "'"
                     .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
                     .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
                     .Formulas(3) = "fxOficina='" & GLOBALES.gOficina & "'"
                     
                     .SelectionFormula = "{REG_CREDITOS.ID_SOLICITUD} = " & vOperacion
                     
                     .SubreportToChange = "sbCorte"
                     .StoredProcParam(0) = vOperacion
                     .StoredProcParam(1) = Format(frmConsultaActiva.dtpCorte.Value, "yyyy/mm/dd")
                
                End If

                .PrintReport
     
               
            End With
            Me.MousePointer = vbDefault
            
      Case 5 'Sep
      Case 6 'Nuevo Credito
                GLOBALES.gCedulaActual = frmConsultaActiva.txtCedula.Text
                Call sbFormsCall("frmCR_SeguimientoTramites")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCR_SeguimientoTramites") Then
                    Call frm.sbGXSegTraIniTlb
                    Exit For
                  End If
                Next frm
      
      Case 7 'Seguimiento de Tramites
            If vOperacion = 0 Then Exit Sub
                
                Call sbFormsCall("frmCR_SeguimientoTramites")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCR_SeguimientoTramites") Then
                    Call frm.sbConsultaExterna(vOperacion)
                    Exit For
                  End If
                Next frm
      
      Case 8 'Sep
      
      Case 9 'Historial
            If vOperacion = 0 Then Exit Sub
                
                Call sbFormsCall("frmCR_ConsultaOperaciones")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCR_ConsultaOperaciones") Then
                    frm.optTipo.Item(0).Value = True
                    Call frm.sbConsultaExterna(vOperacion)
                    Exit For
                  End If
                Next frm
      
      Case 10 'Sep
      Case 11 'Nuevo PreAnalisis
      
            Set x = New clsEstudioCrd
            Set x.vCon = glogon.Conection
            x.xOperacion = vOperacion
            x.xkey = glogon.ConectRPT
      
            x.vSolicitudPreanalisis = 0
            x.vCedula = frmConsultaActiva.txtCedula.Text
    
            Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
            , App.Path, glogon.ConectRPT, 12, glogon.AppName, glogon.AppVersion, glogon.Maquina _
            , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
    
            Set x = Nothing
      
      Case 12 'PreAnalisis
            Set x = New clsEstudioCrd
            Set x.vCon = glogon.Conection
                x.xkey = glogon.ConectRPT
                     
            If .ActiveSheet = 4 Then
                    x.xOperacion = vOperacion
                    x.vSolicitudPreanalisis = vExpediente
                    Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                                , App.Path, glogon.ConectRPT, 11, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                                , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
                    
            Else
                    x.xOperacion = vOperacion
                    strSQL = "select cod_preAnalisis from CRD_PREA_PREANALISIS" _
                           & " Where id_solicitud = " & vOperacion
                    Call OpenRecordSet(rs, strSQL)
                    If rs.EOF And rs.BOF Then
                        Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                        , App.Path, glogon.ConectRPT, 2, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                        , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
                    
                    Else
                        x.vSolicitudPreanalisis = rs!cod_PreAnalisis
                        Call x.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                        , App.Path, glogon.ConectRPT, 11, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                        , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
                    End If
                    rs.Close
                      
            End If
            
      
    
            Set x = Nothing
      
      
      Case 13 'Sep
      
      Case 14 'Plan de Pagos
        If vOperacion = 0 Then Exit Sub
        
        Operacion.OperacionConsulta = vOperacion
        Call sbFormsCall("frmCR_PlanPagos", 1, , , False, Me)
    
      Case 15 'Sep
      
      Case 16 'Cerrar
      
    End Select

End With

Exit Sub
        
vError:
        Me.MousePointer = vbDefault
        MsgBox fxSys_Error_Handler(Err.Description)


End Sub




Private Sub mnuCxCSub_Click(Index As Integer)
Dim strSQL As String, rs As New ADODB.Recordset
Dim vOperacion As Long
Dim frmConsultaActiva As Form, frm As Form

On Error GoTo vError

vOperacion = 0

 
'Localiza el Frm de Consulta que se encuentre activo para utilizarlo de referencia
For Each frmConsultaActiva In Forms
  If (UCase(frmConsultaActiva.Name) = UCase("frmCxC_Consulta")) Then
    'lo localiza y sale (se supone que este procedimiento solo puede ser abierto desde esta pantalla
    Exit For
  End If
Next frmConsultaActiva

With frmConsultaActiva.vgCxC
    .Sheet = .ActiveSheet
    .Row = .ActiveRow
    
    Select Case .Sheet
       Case 1 'Activos
          .col = 2
          If Not IsNumeric(.CellTag) Then Exit Sub
          
          vOperacion = .CellTag
       Case 2, 3 'Cancelados y En Tramite
          .col = 2
          If Not IsNumeric(.Text) Then Exit Sub
          vOperacion = .Text
       
    End Select
  

    Select Case Index
      Case 0 'Abonos
            If vOperacion = 0 Then Exit Sub
      
            .col = 6 'Saldo
             If CCur(.Text) = 0 Then Exit Sub

                    Call sbFormsCall("frmCxC_CuentasAbonos", , , , , Me, True)

                    For Each frm In Forms
                      If (UCase(frm.Name) = UCase("frmCxC_CuentasAbonos")) Then
                        Call frm.sbConsultaExterna(vOperacion)
                        Exit For
                      End If
                    Next frm
    
      
      Case 1 'Anulacion de Abonos
            If vOperacion = 0 Then Exit Sub
                
            Call sbFormsCall("frmCxC_CuentasAnulaciones")
            For Each frm In Forms
              If (UCase(frm.Name) = UCase("frmCxC_CuentasAnulaciones")) Then
                Call frm.sbConsultaExterna(vOperacion)
                Exit For
              End If
            Next frm
      
      Case 2 'Sep
      
     
     
      Case 3 'Nueva Operacion
                GLOBALES.gCedulaActual = frmConsultaActiva.txtCedula.Text
                Call sbFormsCall("frmCxC_Cuentas")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCxC_Cuentas") Then
                    Call frm.sbGXSegTraIniTlb
                    Exit For
                  End If
                Next frm
      
      Case 4 'Tramites
            If vOperacion = 0 Then Exit Sub
                
                Call sbFormsCall("frmCxC_Cuentas")
                For Each frm In Forms
                  If UCase(frm.Name) = UCase("frmCxC_Cuentas") Then
                    Call frm.sbConsultaExterna(vOperacion)
                    Exit For
                  End If
                Next frm
     
      Case 5 'Sep
     
     
      Case 6 'Movimientos de la Operacion
            
            If vOperacion = 0 Then Exit Sub
            
            Me.MousePointer = vbHourglass
            
            With frmContenedor.Crt
                .Reset
                .WindowShowPrintSetupBtn = True
                .WindowShowRefreshBtn = True
                .WindowShowSearchBtn = True
                .WindowState = crptMaximized
                .WindowTitle = "Reportes del Módulo de CxC"
                
                .Connect = glogon.ConectRPT
                
                     .ReportFileName = SIFGlobal.fxPathReportes("CxC_PlanPagosMov.rpt")
                    
                     .Formulas(0) = "fxFecha='FECHA: " & Format(fxFechaServidor, "dd/mm/yyyy  hh:mm:ss") & "'"
                     .Formulas(1) = "fxEmpresa='" & GLOBALES.gstrNombreEmpresa & "'"
                     .Formulas(2) = "fxUsuario='USER: " & glogon.Usuario & "'"
                     .Formulas(3) = "fxOficina='" & GLOBALES.gOficina & "'"
                     
                     .SelectionFormula = "{CXC_CUENTAS.OPERACION} = " & vOperacion
                     
'                     .SubreportToChange = "sbCorte"
'                     .StoredProcParam(0) = vOperacion
'                     .StoredProcParam(1) = Format(frmConsultaActiva.dtpCorte.Value, "yyyy/mm/dd")
                .PrintReport

               
            End With
            Me.MousePointer = vbDefault
            
      Case 7 'Plan de Pagos
        If vOperacion = 0 Then Exit Sub
        
        Operacion.OperacionConsulta = vOperacion
        Call sbFormsCall("frmCxC_PlanPagos", 1, , , False, Me)
      
      Case 5 'Sep
      
      Case 8 'Sep
      
      Case 9 'Cerrar
      
    End Select

End With

Exit Sub
        
vError:
        Me.MousePointer = vbDefault
        MsgBox fxSys_Error_Handler(Err.Description)


End Sub




Private Sub mnuAyudaAcercaDe_Click()
 frmAcercaDe.Show vbModal
End Sub

Private Sub mnuAyudaContenido_Click()
   frmContenedor.CD.HelpCommand = cdlHelpContents
   frmContenedor.CD.ShowHelp
   frmContenedor.CD.HelpCommand = cdlHelpContext
End Sub



Private Sub mnuParametrosSistema_Click(Index As Integer)
Dim Nucleo As clsNucleo
  
Set Nucleo = New clsNucleo
  
Select Case Index
  Case 0 'Cofig. Empresa
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 2, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
  
  Case 1 'Comunicados de Servicio
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 3, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
  
  Case 2 'Separador
  
  Case 3 'Encabezados y Pie de Pagina EC
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 4, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
  
  Case 4 'Separador
  
  Case 5 'Consulta Cola de Asientos
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 6, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
  
  Case 6 'Separador
  
  Case 7 'Oficinas
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 7, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
  Case 8 'Oficinas Metas
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 8, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
  Case 9 'Separador
  
  Case 10 'Varibales Globales
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 1, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)

End Select

Set Nucleo = Nothing

End Sub

Private Sub mnuPE_Modo_1_Click()
Call sbFormsCall("frmCC_PE_PlanillaDirecta", , , , False)
End Sub

Private Sub mnuSalir_Click()
On Error Resume Next
 
  'ODBC: [CORE] Crystal Reports
 glogon.DSN = "PGX_Core"
 Call sbLogonDSN(glogon.DSN, True, 0)
   
 'ODBC: [PORTAL] Crystal Reports
 glogon.DSN = "PGX_Portal"
 Call sbLogonDSN(glogon.DSN, True, 1)
 
 'ODBC: [AUXILIARES] Crystal Reports
 glogon.DSN = "PGX_Auxiliar"
 Call sbLogonDSN(glogon.DSN, True, 2)
 
 'ODBC: [AUXILIARES] Crystal Reports
 glogon.DSN = "PGX_Analisis"
 Call sbLogonDSN(glogon.DSN, True, 3)
 
 
 Call sbSEGCuentaLog("11")
' glogon.Conection.Close
 End
End Sub


Private Sub MDIForm_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Shift = 1 And Button = 2 Then MsgBox App.FileDescription & vbCrLf & App.LegalCopyright
End Sub

Private Sub mnuSeguridadSub_Click(Index As Integer)
Dim frmX As Form, pUsuario As String


Select Case Index

 Case 0 'Cambia Contraseña
         frmCambiaClave.Show vbModal
 
 Case 1 'Actualiza Datos de Contacto de Usuario
         frmLogon_Datos_Update.Show vbModal
 
 Case 2 'Sep
 
 Case 3 'Cambiar de Tema
         frmLogon_Theme.Show vbModal
 
 
 Case 4 'Sep
 
 Case 5 'Bitacora
        Dim Nucleo As clsNucleo
        Set Nucleo = New clsNucleo
        
        Call Nucleo.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 5, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
        
        Set Nucleo = Nothing

End Select

End Sub


Private Sub mnuVerSub_Click(Index As Integer)
Dim frmX As Form

Select Case Index
  Case 0 'Ordenar x Iconos
     Me.Arrange vbArrangeIcons
  Case 1 'Ordenar en Cascada
     Me.Arrange vbCascade
  Case 2 'Ordenar x Titulo Vertical
     Me.Arrange vbTileVertical
  Case 3 'Ordenar x Titulo Horizonal
     Me.Arrange vbTileHorizontal
  
  Case 4 'Separador
  
  Case 5 'Cerrar todas las ventanas
     For Each frmX In Forms
      If Not (frmX Is Me) Then
         Unload frmX
      End If
     Next frmX
   
  Case 6 'Minimizar todas las ventanas
     For Each frmX In Forms
      If Not (frmX Is Me) Then
         frmX.WindowState = vbMinimized
      End If
     Next frmX
   
  Case 7 'Restaurar todas las ventanas
     For Each frmX In Forms
      If Not (frmX Is Me) Then
         frmX.WindowState = vbNormal
      End If
     Next frmX
   

End Select
End Sub

Private Sub sbLogonReconexion()
'Verifica que la conexión se encuentre activa

'Dim i As Integer, vMenu As String
'
'On Error GoTo vError
'
'If glogon.Reconexion = 5 Then Exit Sub
'
'vMenu = Me.Caption
'
'glogon.Conection.CommandTimeout = 10
'
'If glogon.Conection.State = 1 Then glogon.Conection.Close
'
'Me.Caption = "Conexión Caída..Reintentando Conectar al Servidor..(" & glogon.Reconexion & ")"
'glogon.Conection.Open
'
'glogon.Conection.CommandTimeout = 360
'glogon.Reconexion = 1
'Me.Caption = vMenu
'MsgBox "Conexión Reestablecida!", vbInformation
'
'Exit Sub
'
'vError:
' If glogon.Reconexion < 5 And glogon.Conection.State = 0 Then
'       glogon.Reconexion = glogon.Reconexion + 1
'       Me.Caption = vMenu
'       MsgBox "No fue posible la conexión con el servidor...intente nuevamenta la reconexión", vbCritical
' End If
'
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

'Barra de Favoritos
If pLoad Then
        Call sbSIFFavoritosToolBar(tlbFavoritos)
End If

CoolBarX.Bands.Item(2).MinWidth = 2000
CoolBarX.Bands.Item(3).MinWidth = 900
CoolBarX.Bands.Item(5).Caption = gPortal.Empresa_Name
CoolBarX.Bands.Item(5).BackColor = RGB(36, 113, 163)


'Revisa si el usuario cumple con requisitos de accesos del Cliente
strSQL = "exec spSEG_Access_Limit " & gPortal.Empresa_Id & ",'" & glogon.Usuario & "','" & glogon.Maquina & "',''"
Call OpenRecordSet(rs, strSQL, 1)

If rs!Indicador = 0 Then '(1 Pasa, 0 No Pasa)
       
   Select Case rs!Indicador
     Case -1
            MsgBox "Su estación de trabajo ha sido desvinculada!", vbExclamation
     Case -2
            MsgBox "Lo sentimos su sesión de Trabajo a Expirado!", vbExclamation
   End Select
   
   'Termina Sesión
   Call mnuSalir_Click

End If
rs.Close




strSQL = "select Fecha_Congela,isnull(fecha_Congela,getdate()) as 'Fecha_Auxiliar', Getdate() as 'Fecha_Actual'" _
        & " from SIF_Empresa"
Call OpenRecordSet(rs, strSQL)

statusBar.Pane(1).Text = glogon.Usuario
statusBar.Pane(2).Text = "Fecha Auxiliar: " & Format(rs!Fecha_Auxiliar, "dd/mm/yyyy")
statusBar.Pane(3).Text = "Fecha Actual: " & Format(rs!Fecha_Actual, "dd/mm/yyyy")

'statusBar.Panels.Item(4).Text = glogon.Usuario
'statusBar.Panels.Item(5).Text = "Fecha Auxiliar: " & Format(rs!Fecha_Auxiliar, "dd/mm/yyyy")
'statusBar.Panels.Item(6).Text = "Fecha Actual: " & Format(rs!Fecha_Actual, "dd/mm/yyyy")

If IsNull(rs!Fecha_Congela) Then
    CoolBarX.Bands.Item(2).Visible = False
Else
    CoolBarX.Bands.Item(2).Visible = True
    CoolBarX.Bands.Item(2).Caption = "Fecha Bloqueada: " & Format(rs!Fecha_Auxiliar, "dd/mm/yyyy")
    CoolBarX.Bands.Item(2).Width = 3200
End If

rs.Close


If mLoad_Inicial Then
        strSQL = "exec  spSEG_Logon_Info '" & glogon.Usuario & "','" & glogon.Maquina_MAC & "'"
        Call OpenRecordSet(rs, strSQL, 1)
        
     If Len(Trim(rs!Tel_Cell & "")) = 0 Or Len(Trim(rs!Email & "")) = 0 Then
         frmLogon_Datos_Update.Show vbModal
     End If
    rs.Close
End If
 
mLoad_Inicial = False

Exit Sub

vErrorTimerLoad:
  MsgBox fxSys_Error_Handler(Err.Description), vbCritical
End Sub

Private Sub TimerSalir_Timer()
TimerSalir.Interval = 0
Call mnuSalir_Click
End Sub


Private Sub tlbMenu_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
Dim clsMarcas As clsMarcas

Set clsMarcas = New clsMarcas
 
Select Case ButtonMenu.Key
 Case "Marca"
        Call clsMarcas.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
            , App.Path, glogon.ConectRPT, 1, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
 
 Case "MarcasDetalle"
        Call clsMarcas.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
            , App.Path, glogon.ConectRPT, 4, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
 Case "Configuracion"
        Call clsMarcas.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
            , App.Path, glogon.ConectRPT, 3, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)

 Case "AsgUsuarios"
        Call clsMarcas.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
            , App.Path, glogon.ConectRPT, 2, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)

End Select

Set clsMarcas = Nothing

End Sub

Private Sub tlbFavoritos_ButtonClick(ByVal Button As MSComctlLib.Button)

Call sbSIFMenuOptionClick(Button.Tag)

End Sub

Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)

Dim i As Integer, frmX As Form


Call sbFormActivo("frmCntX_Explorer", frmX)
If Not frmX Is Nothing Then
   frmX.Hide
Else
End If

Call sbFormActivo("frmActivos_Explorador", frmX)
If Not frmX Is Nothing Then
   frmX.Hide
Else
End If

Select Case Button.Key
    Case "Menu"

        Call sbFormsCall("frmMenu", 0, 1, 1)
        
    Case "Contabilidad"
        
'        tlbPeriodo.Width = 1080
'        txtAnio.Left = tlbPeriodo.Left + 1130
'        txtMes.Left = txtAnio.Left + txtAnio.Width + 20
'
'        tlbPeriodo.Top = 420
'
'        txtAnio.Top = tlbPeriodo.Top
'        txtMes.Top = tlbPeriodo.Top
'        lblPeriodo.Top = tlbPeriodo.Top + 40
'        lblPeriodo.Left = txtMes.Left + txtMes.Width + 60
        
        If Not CoolBarContabilidad.Visible Then
            CoolBarContabilidad.Visible = True
            
            txtAnio.Visible = True
            txtMes.Visible = True
            lblPeriodo.Visible = True
        Else
            CoolBarContabilidad.Visible = False
        
            txtAnio.Visible = False
            txtMes.Visible = False
            lblPeriodo.Visible = False
        End If
        
        Call sbFormsCall("frmCntX_Explorer")
        Call sbFormActivo("frmCntX_Explorer", frmX)
        If Not frmX Is Nothing Then
          frmX.WindowState = vbMaximized
        Else
        End If

        
    Case "Activos"
        Call sbFormsCall("frmActivos_Explorador")
        Call sbFormActivo("frmActivos_Explorador", frmX)
        If Not frmX Is Nothing Then
          frmX.WindowState = vbMaximized
        Else
        End If
  Case "Impresoras"
    Call sbFormsCall("frmCC_Impresoras")
  
  Case "Marcas"
        Dim Marcas As clsMarcas

        Set Marcas = New clsMarcas
                Call Marcas.Menu(glogon.Usuario, glogon.Conection, glogon.BaseDatos, glogon.Servidor _
                    , App.Path, glogon.ConectRPT, 1, glogon.AppName, glogon.AppVersion, glogon.Maquina _
                    , glogon.Portal_Con, glogon.Portal_User, glogon.Portal_Key, gPortal.Empresa_Id, gPortal.Empresa_Name)
        Set Marcas = Nothing
End Select


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
     ' Call sbCntX_RestructuraMovimientosRSM(MDIMenu.txtAnio, MDIMenu.txtMes, frmX, False)
      Call sbCntX_CierreFiscal(frmX, txtMes.Text, txtAnio.Text)
    End If

  Case "Balance" 'Revision del Balance
    
    iRespuesta = MsgBox("Esta seguro que desea revisar la balanza de comprobación por inconsistencias?", vbYesNo)
    If iRespuesta = vbYes Then
       Set frmX = frmCntX_Procesos
       Call sbCntX_RestructuraMovimientosRSM(gCntX_Parametros.PeriodoAnio, gCntX_Parametros.PeriodoMes, frmX)
    End If
  
  Case "Asientos" 'Verificacion de Asientos
    Call sbFormsCall("frmCntX_UtilVerificaAsientos", vbModal, , , False, Me)

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

tlbPeriodo.top = txtAnio.top - 40
tlbPeriodo.Left = tlbPeriodo.Left - 60
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

Private Sub sbCntX_Periodo_Refresh()
Dim strSQL As String, rs As New ADODB.Recordset
Dim strResultado As String

On Error GoTo vError

txtAnio = Val(txtAnio)
  
  
gActivos.Anio = txtAnio.Text
gActivos.Mes = txtMes.Text
gActivos.Periodo = CDate(txtAnio.Text & "/" & Format(txtMes.Text, "00") & "/01")
gActivos.Periodo = DateAdd("d", -1, DateAdd("m", 1, gActivos.Periodo))

  
lblPeriodo.Text = fxCntX_PeriodoDesc(txtAnio, txtMes)
'Call frmCntX_Explorer.sbRefrescaArbol


strSQL = "select estado from CntX_Periodos where cod_contabilidad = " & gCntX_Parametros.CodigoConta _
       & " and anio = " & txtAnio & " and mes = " & txtMes
Call OpenRecordSet(rs, strSQL, 0)
If rs.EOF And rs.BOF Then
 tlbPeriodo.Buttons.Item(1).ToolTipText = "Periodo No Definido"
 tlbPeriodo.Buttons.Item(1).Image = 5
 lblPeriodo.ForeColor = vbRed
Else
  If rs!Estado = "P" Then
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
  
tlbPeriodo.top = 0
Exit Sub

vError:

End Sub

